#!/usr/bin/env python3
"""Run the pure-text VBA unit tests headless in LibreOffice Basic.

LibreOffice Basic with `Option VBASupport 1` understands the VBA used in the
Word-independent modules (modStrings, modClaimParse, modJson, modDefaults).
This lets the parsing logic be tested without Word:
  - tests/vba/*.bas unit tests
  - every tests/fixtures/*.docx: its visible text is run through PlanRenumber
    and compared with <name>.expected.txt (from tools/make_fixtures.py)
The Word layer (modRenumber, ...) still has to be tested in Word itself.

Usage:  python3 tools/run_vba_tests.py
Exit code 0 = all tests passed.
"""
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MODULES = [
    ROOT / "src/modules/modStrings.bas",
    ROOT / "src/modules/modClaimParse.bas",
    ROOT / "src/modules/modJson.bas",
    ROOT / "src/modules/modDefaults.bas",
    ROOT / "tests/vba/modTestRenumber.bas",
    ROOT / "tests/vba/modTestJson.bas",
]
TEST_FUNCTIONS = ["RunRenumberTests", "RunJsonTests", "RunFixtureTests"]
FIXTURES = ROOT / "tests" / "fixtures"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Scripting.Dictionary exists only on Windows; LibreOffice gets this stand-in.
DICT_SHIM = """Option VBASupport 1
Option ClassModule
Option Explicit
Private mKeys() As String
Private mVals() As Variant
Private mCount As Long

Private Function IndexOf(ByVal k As String) As Long
    Dim i As Long
    IndexOf = -1
    For i = 0 To mCount - 1
        If StrComp(mKeys(i), k, 0) = 0 Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function Exists(ByVal k As String) As Boolean
    Exists = (IndexOf(k) >= 0)
End Function

Public Sub Add(ByVal k As String, ByVal v As Variant)
    ReDim Preserve mKeys(0 To mCount)
    ReDim Preserve mVals(0 To mCount)
    mKeys(mCount) = k
    If IsObject(v) Then
        Set mVals(mCount) = v
    Else
        mVals(mCount) = v
    End If
    mCount = mCount + 1
End Sub

Public Sub Remove(ByVal k As String)
    Dim i As Long, j As Long
    i = IndexOf(k)
    If i < 0 Then Exit Sub
    For j = i To mCount - 2
        mKeys(j) = mKeys(j + 1)
        If IsObject(mVals(j + 1)) Then
            Set mVals(j) = mVals(j + 1)
        Else
            mVals(j) = mVals(j + 1)
        End If
    Next j
    mCount = mCount - 1
End Sub

Public Function Item(ByVal k As String) As Variant
    Dim i As Long
    i = IndexOf(k)
    If i < 0 Then Err.Raise 5
    If IsObject(mVals(i)) Then
        Set Item = mVals(i)
    Else
        Item = mVals(i)
    End If
End Function
"""
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def to_xba(name: str, code: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
        f'<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{name}" '
        f'script:language="StarBasic" script:moduleType="normal">{html.escape(code, quote=False)}</script:module>\n'
    )


def basic_source(path: Path) -> str:
    code = path.read_text(encoding="ascii").replace("\r\n", "\n")
    code = re.sub(r"^Attribute .*\n", "", code, flags=re.M)
    code = code.replace('CreateObject("Scripting.Dictionary")', "New PTDict")
    code = re.sub(r"^\s*d\.CompareMode = .*\n", "", code, flags=re.M)
    return "Option VBASupport 1\n" + code


def visible_paragraphs(docx: Path) -> list[str]:
    """Paragraph texts as Word shows them with tracked deletions removed."""
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    result = []
    for p in root.iter(W + "p"):
        parts = []
        for el in p.iter():
            if el.tag == W + "t":
                parts.append(el.text or "")
            elif el.tag == W + "tab":
                parts.append("\t")
        result.append("".join(parts))
    return result


def basic_literal(text: str) -> str:
    """Basic expression for text using only ASCII source characters."""
    pieces, buf = [], ""
    for ch in text:
        if 32 <= ord(ch) < 127:
            buf += '""' if ch == '"' else ch
        else:
            if buf:
                pieces.append(f'"{buf}"')
                buf = ""
            pieces.append("vbTab" if ch == "\t" else f"ChrW({ord(ch)})")
    if buf or not pieces:
        pieces.append(f'"{buf}"')
    return " & ".join(pieces)


def fixture_module() -> str:
    """Generates RunFixtureTests: feeds every fixture's visible text through
    PlanRenumber and compares with <name>.expected.txt."""
    lines = ["Option VBASupport 1", "Option Explicit", "",
             "Public Function RunFixtureTests() As String",
             "    Dim r As String, failed As Long, passed As Long",
             "    Dim inp() As String, exp() As String",
             "    Dim ops As Collection, errMsg As String, n As Long, outp As Variant",
             "    Dim flagLog As String, flags As Long, i As Long, ok As Boolean"]
    for docx in sorted(FIXTURES.glob("*.docx")):
        exp_lines = docx.with_suffix(".expected.txt").read_text(encoding="utf-8").rstrip("\n").split("\n")
        exp_flags = int(exp_lines[0].split()[1])
        exp_pars = exp_lines[1:]
        inp_pars = visible_paragraphs(docx)
        lines.append(f"    ' {docx.name}")
        lines.append(f"    ReDim inp(0 To {len(inp_pars) - 1})")
        lines += [f"    inp({i}) = {basic_literal(t)}" for i, t in enumerate(inp_pars)]
        lines.append(f"    ReDim exp(0 To {len(exp_pars) - 1})")
        lines += [f"    exp({i}) = {basic_literal(t)}" for i, t in enumerate(exp_pars)]
        lines += [
            "    Set ops = New Collection",
            "    ok = PlanRenumber(inp, ops, errMsg, n)",
            "    If ok Then",
            "        outp = ApplyOpsToTexts(inp, ops, flagLog)",
            "        flags = 0",
            "        If Len(flagLog) > 0 Then flags = UBound(Split(flagLog, vbLf))",
            f"        If flags <> {exp_flags} Then",
            "            ok = False",
            f'            r = r & "FAIL {docx.stem}: flags " & flags & " expected {exp_flags}" & vbLf & flagLog',
            "        End If",
            f"        If UBound(outp) <> {len(exp_pars) - 1} Then",
            "            ok = False",
            f'            r = r & "FAIL {docx.stem}: paragraph count" & vbLf',
            "        Else",
            "            For i = 0 To UBound(exp)",
            "                If outp(i) <> exp(i) Then",
            "                    ok = False",
            f'                    r = r & "FAIL {docx.stem} paragraph " & i & ": [" & outp(i) & "] expected [" & exp(i) & "]" & vbLf',
            "                End If",
            "            Next i",
            "        End If",
            "    Else",
            f'        r = r & "FAIL {docx.stem}: " & errMsg & vbLf',
            "    End If",
            "    If ok Then",
            "        passed = passed + 1",
            "    Else",
            "        failed = failed + 1",
            "    End If",
        ]
    lines += ['    RunFixtureTests = "Fixtures passed: " & passed & "  Failed: " & failed & vbLf & r',
              "End Function", ""]
    return "\n".join(lines)


def main() -> int:
    if not SOFFICE:
        print("LibreOffice (soffice) not found", file=sys.stderr)
        return 2
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        profile = tmp / "profile"
        env_arg = f"-env:UserInstallation={profile.as_uri()}"
        subprocess.run([SOFFICE, env_arg, "--headless", "--terminate_after_init"],
                       check=True, capture_output=True, timeout=120)
        lib = profile / "user/basic/Standard"
        lib.mkdir(parents=True, exist_ok=True)

        out_file = tmp / "result.txt"
        runner = (
            "Option VBASupport 1\nOption Explicit\n"
            "Sub Main\n"
            "    Dim f As Integer\n"
            "    f = FreeFile\n"
            f'    Open "{out_file}" For Output As #f\n'
            + "".join(f"    Print #f, {fn}()\n" for fn in TEST_FUNCTIONS) +
            "    Close #f\n"
            "End Sub\n"
        )
        names = []
        for path in MODULES:
            names.append(path.stem)
            (lib / f"{path.stem}.xba").write_text(to_xba(path.stem, basic_source(path)), encoding="utf-8")
        names.append("modFixtureTests")
        (lib / "modFixtureTests.xba").write_text(to_xba("modFixtureTests", fixture_module()), encoding="utf-8")
        names.append("PTDict")
        (lib / "PTDict.xba").write_text(to_xba("PTDict", DICT_SHIM), encoding="utf-8")
        names.append("PTRunner")
        (lib / "PTRunner.xba").write_text(to_xba("PTRunner", runner), encoding="utf-8")
        elements = "\n".join(f' <library:element library:name="{n}"/>' for n in names)
        (lib / "script.xlb").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
            '<library:library xmlns:library="http://openoffice.org/2000/library" '
            'library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
            f"{elements}\n</library:library>\n",
            encoding="utf-8",
        )

        url = "vnd.sun.star.script:Standard.PTRunner.Main?language=Basic&location=application"
        proc = subprocess.run([SOFFICE, env_arg, "--headless", "--norestore", url],
                              capture_output=True, text=True, timeout=300)
        if not out_file.exists():
            print("No test output. LibreOffice said:\n", proc.stdout, proc.stderr, file=sys.stderr)
            return 2
        result = out_file.read_text(encoding="utf-8", errors="replace")
        print(result)
        ok = len(re.findall(r"Failed: 0\b", result)) == len(TEST_FUNCTIONS)
        return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
