#!/usr/bin/env python3
"""Static checks for the VBA sources (no Office needed).

Catches what would otherwise only show up as a compile error in Word
("Variable not defined", "Sub or Function not defined", "Ambiguous name",
"Block If without End If", ...):

  1. every identifier is declared: local / parameter / module-level / public in
     another module / VBA or Word object model name (allow-list below)
  2. block structure: If/For/Do/While/With/Select/Sub/Function balance
  3. public names are unique across modules; no duplicate procedures in a module

Usage: python3 tools/vba_lint.py [files...]   (default: src/ tests/vba/ tools/*.bas)
Exit code 0 = no problems.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def rel(path):
    try:
        return path.relative_to(ROOT)
    except ValueError:
        return path

KEYWORDS = set("""
and as boolean byref byte byval call case const currency date debug declare dim do double each else elseif empty end
enum erase error exit explicit false for friend function get global gosub goto if implements in integer is let lib like
long loop me mod new next not nothing null object on option optional or paramarray preserve private property public
raiseevent redim rem resume return select set single static step stop string sub then to true type typeof until variant
wend while with withevents xor attribute vb_name compare binary text base module private
""".split())

# VBA functions/statements and objects used in this project
VBA_BUILTINS = set("""
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chr$ chrw chrw$ cint clng clnglng cos createobject csng cstr
cvar date$ dateadd datediff datepart dateserial day dir dir$ doevents environ environ$ eof err error$ exp filecopy
filedatetime filelen fix format format$ freefile getattr hex hex$ hour iif instr instrrev int isarray isdate isempty
iserror ismissing isnull isnumeric isobject join kill lbound lcase lcase$ left left$ len lenb loadpicture log ltrim
ltrim$ mid mid$ minute mkdir month msgbox now oct oct$ replace right right$ rmdir rnd round rtrim rtrim$ second seek
sgn shell sin space space$ split sqr str str$ strcomp strconv string$ strreverse tab tan time time$ timer timeserial
timevalue trim trim$ typename ubound ucase ucase$ val vartype weekday year inputbox erl collection
vbcr vblf vbcrlf vbtab vbnullstring vbnullchar vbnewline vbobjecterror vbtextcompare vbbinarycompare
vbokonly vbokcancel vbyesno vbyesnocancel vbyes vbno vbcancel vbok vbinformation vbexclamation vbcritical vbquestion
vbdirectory vbnormal vbnormalfocus vbhidden vbsystem vbreadonly vbarchive
open close print output append input line binary random access write read shared lock
""".split())

# Word / Office object model names used unqualified (globals, types, constants)
WORD_NAMES = set("""
application activedocument documents selection document range paragraph paragraphs paragraphformat font comment
comments revision revisions section sections headerfooter view table iribboncontrol iribbonui
wdyellow wdbrightgreen wdturquoise wdpink wdblue wdred wddarkblue wdteal wdgreen wdviolet wddarkred wddarkyellow
wdgray50 wdgray25 wdnohighlight wdrevisiondelete wdrevisionmovedfrom wdrevisionsmarkupall wdrevisionsviewfinal
wdinlinerevisions wdlinespacesingle wdlinespace1pt5 wdlinespacedouble wdlinespacemultiple wdseparatebytabs
wdheaderfooterprimary wdheaderfooterfirstpage wdheaderfooterevenpages wdcharacter wdalignparagraphleft
wdaligntabright wdmaintextstory wdcollapseend wdfindstop wdreplaceall wddonotsavechanges
msofiledialogfolderpicker msofiledialogfilepicker
linestopoints centimeterstopoints
""".split())

PROC_RE = re.compile(r"^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*(?:\((.*)\))?", re.I)
DECL_RE = re.compile(r"^(?:(Public|Private|Global|Dim|Static)\s+)(?:WithEvents\s+)?(.*)$", re.I)
CONST_RE = re.compile(r"^(?:(Public|Private|Global)\s+)?Const\s+(\w+)", re.I)
IDENT_RE = re.compile(r"(?<![\w.$#&])([A-Za-z_]\w*\$?)")


def logical_lines(text):
    """(line number, code) with comments and string contents removed, continuations joined."""
    out, buf, start = [], "", None
    for no, raw in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        code, in_str, i = [], False, 0
        while i < len(raw):
            ch = raw[i]
            if ch == '"':
                if in_str and i + 1 < len(raw) and raw[i + 1] == '"':
                    i += 2
                    continue
                in_str = not in_str
                code.append('"')
            elif not in_str and ch == "'":
                break
            elif not in_str:
                code.append(ch)
            i += 1
        line = "".join(code).rstrip()
        if start is None:
            start = no
        if line.endswith(" _"):
            buf += line[:-2] + " "
            continue
        out.append((start, (buf + line).strip()))
        buf, start = "", None
    return out


def split_decl_names(rest):
    """Names from 'a As Long, b() As String, c' (array bounds and types removed)."""
    rest = re.sub(r"\([^()]*\)", "", rest)
    names = []
    for part in rest.split(","):
        part = part.strip()
        m = re.match(r"(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?(\w+)", part, re.I)
        if m:
            names.append(m.group(1).lower())
    return names


def module_info(path):
    text = path.read_text(encoding="ascii", errors="replace")
    lines = logical_lines(text)
    public, private, procs = set(), set(), []
    in_proc = False
    for no, code in lines:
        if not code:
            continue
        m = PROC_RE.match(code)
        if m:
            name = m.group(3).lower()
            (private if (m.group(1) or "").lower() == "private" else public).add(name)
            procs.append(name)
            in_proc = True
            continue
        if re.match(r"^End\s+(Sub|Function|Property)\b", code, re.I):
            in_proc = False
            continue
        if in_proc:
            continue
        m = CONST_RE.match(code)
        if m:
            (private if (m.group(1) or "").lower() == "private" else public).add(m.group(2).lower())
            continue
        m = DECL_RE.match(code)
        if m and not re.match(r"^(Public|Private)\s+(Declare|Type|Enum)\b", code, re.I):
            target = private if m.group(1).lower() in ("private", "dim") else public
            target.update(split_decl_names(m.group(2)))
    return text, lines, public, private, procs


def check_blocks(path, lines):
    problems, stack = [], []
    for no, code in lines:
        c = code
        if not c:
            continue
        lc = c.lower()
        if re.match(r"^\w+:$", c):          # label
            continue
        # single-line If: something after Then
        m = re.match(r"^(?:else)?if\b.*\bthen\b(.*)$", lc)
        if lc.startswith("if ") or lc.startswith("elseif "):
            after = m.group(1).strip() if m else ""
            if lc.startswith("if ") and after == "":
                stack.append(("if", no))
            elif lc.startswith("elseif ") and (not stack or stack[-1][0] != "if"):
                problems.append(f"{path}:{no}: ElseIf without If")
            continue
        opener = None
        if re.match(r"^for\b", lc):
            opener = "for"
        elif re.match(r"^do\b", lc):
            opener = "do"
        elif re.match(r"^while\b", lc):
            opener = "while"
        elif re.match(r"^with\b", lc):
            opener = "with"
        elif re.match(r"^select\s+case\b", lc):
            opener = "select"
        elif PROC_RE.match(c):
            opener = "proc"
        if opener:
            stack.append((opener, no))
            continue
        closers = [(r"^end\s+if\b", "if"), (r"^next\b", "for"), (r"^loop\b", "do"), (r"^wend\b", "while"),
                   (r"^end\s+with\b", "with"), (r"^end\s+select\b", "select"),
                   (r"^end\s+(sub|function|property)\b", "proc")]
        for pat, kind in closers:
            if re.match(pat, lc):
                if not stack or stack[-1][0] != kind:
                    top = f"{stack[-1][0]} from line {stack[-1][1]}" if stack else "nothing"
                    problems.append(f"{path}:{no}: '{code}' closes {top}")
                else:
                    stack.pop()
                break
        if lc == "else" and (not stack or stack[-1][0] != "if"):
            problems.append(f"{path}:{no}: Else without If")
    for kind, no in stack:
        problems.append(f"{path}:{no}: {kind} block not closed")
    return problems


def check_identifiers(path, lines, known_global, module_private):
    problems = []
    local = set()
    in_proc = False
    for no, code in lines:
        if not code:
            continue
        m = PROC_RE.match(code)
        if m:
            in_proc = True
            local = set(split_decl_names(m.group(4) or ""))
            local.add(m.group(3).lower())
            continue
        if re.match(r"^End\s+(Sub|Function|Property)\b", code, re.I):
            in_proc = False
            continue
        if not in_proc:
            continue
        lm = re.match(r"^(\w+):$", code)
        if lm:
            local.add(lm.group(1).lower())
            continue
        dm = re.match(r"^(?:Dim|Static|ReDim(?:\s+Preserve)?)\s+(.*)$", code, re.I)
        if dm:
            local.update(split_decl_names(dm.group(1)))
        cm = re.match(r"^Const\s+(\w+)", code, re.I)
        if cm:
            local.add(cm.group(1).lower())
    # second pass: uses
    local, in_proc = set(), False
    for no, code in lines:
        if not code:
            continue
        m = PROC_RE.match(code)
        if m:
            in_proc = True
            local = set(split_decl_names(m.group(4) or ""))
            local.add(m.group(3).lower())
            # collect all locals of this procedure up front (Dim may come later)
            continue
        if re.match(r"^End\s+(Sub|Function|Property)\b", code, re.I):
            in_proc = False
            continue
        if not in_proc:
            continue
        lm = re.match(r"^(\w+):$", code)
        if lm:
            local.add(lm.group(1).lower())
            continue
        dm = re.match(r"^(?:Dim|Static|ReDim(?:\s+Preserve)?)\s+(.*)$", code, re.I)
        if dm:
            local.update(split_decl_names(dm.group(1)))
            # types after "As" are checked below via the general scan, skip declaration lines
            check = re.sub(r"\bAs\s+(New\s+)?\w+", "", dm.group(1))
            check = re.sub(r"^\s*\w+|,\s*\w+", ",", check)
            code_to_scan = check
        else:
            code_to_scan = code
        cm = re.match(r"^Const\s+(\w+)\s+As\s+\w+\s*=(.*)$", code, re.I)
        if cm:
            local.add(cm.group(1).lower())
            code_to_scan = cm.group(2)
        # remove string literals, named arguments, type names, hex/number literals
        scan = re.sub(r'"[^"]*"', '""', code_to_scan)
        scan = re.sub(r"\b\w+:=", "", scan)
        scan = re.sub(r"\bAs\s+(New\s+)?\w+", "", scan, flags=re.I)
        scan = re.sub(r"\bNew\s+\w+", "New", scan, flags=re.I)
        scan = re.sub(r"&H[0-9A-Fa-f]+&?", "0", scan)
        scan = re.sub(r"\bGoTo\s+\w+", "", scan, flags=re.I)
        scan = re.sub(r"\bOn Error\s+(GoTo\s+\w+|Resume Next|GoTo 0)", "", scan, flags=re.I)
        for ident in IDENT_RE.findall(scan):
            low = ident.lower()
            if low in KEYWORDS or low in VBA_BUILTINS or low in WORD_NAMES:
                continue
            if low.rstrip("$") in VBA_BUILTINS:
                continue
            if low in local or low in module_private or low in known_global:
                continue
            problems.append(f"{path}:{no}: '{ident}' is not declared (or not public in any module)")
    return problems


RESERVED_AS_NAMES = set("""name date time error input open close print line get put seek lock width write string type
step stop end exit len mid left right format val kill beep load unload reset randomize erase let set call""".split())


def check_physical(path, text):
    """VBA limits: 1023 characters per physical line, 24 continuations per statement."""
    problems, cont = [], 0
    for no, raw in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        if len(raw) > 1023:
            problems.append(f"{path}:{no}: line longer than 1023 characters")
        if raw.rstrip().endswith(" _"):
            cont += 1
            if cont > 24:
                problems.append(f"{path}:{no}: more than 24 line continuations")
        else:
            cont = 0
    return problems


def check_reserved_names(path, lines):
    problems = []
    for no, code in lines:
        names = []
        m = PROC_RE.match(code)
        if m:
            names = split_decl_names(m.group(4) or "")
        dm = re.match(r"^(?:Dim|Static|Private|Public)\s+(?!Sub\b|Function\b|Const\b|Property\b|Declare\b)(.*)$", code, re.I)
        if dm:
            names = split_decl_names(dm.group(1))
        for n in names:
            if n in RESERVED_AS_NAMES:
                problems.append(f"{path}:{no}: '{n}' is a VBA statement/function name - rename the variable")
    return problems


PARAM_RE = re.compile(r"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?(\w+)(\(\))?(?:\s+As\s+(\w+))?", re.I)


def proc_signatures(files_lines):
    """name -> list of (mode, type, isArray) for every parameter."""
    sigs = {}
    for lines in files_lines:
        for no, code in lines:
            m = PROC_RE.match(code)
            if not m:
                continue
            params = []
            for part in [x.strip() for x in (m.group(4) or "").split(",") if x.strip()]:
                pm = PARAM_RE.match(part)
                if pm:
                    mode = "byval" if (pm.group(2) or "").strip().lower() == "byval" else "byref"
                    kind = "optional" if pm.group(1) else ("paramarray" if pm.group(3) else "required")
                    params.append((mode, (pm.group(6) or "variant").lower(), bool(pm.group(5)), kind))
            sigs[m.group(3).lower()] = params
    return sigs


def check_byref_types(path, lines, sigs):
    """A variable passed to a ByRef parameter must have exactly the parameter's type."""
    problems, types = [], {}
    for no, code in lines:
        if PROC_RE.match(code):
            types = {}
            m = PROC_RE.match(code)
            for part in [x.strip() for x in (m.group(4) or "").split(",") if x.strip()]:
                pm = PARAM_RE.match(part)
                if pm:
                    types[pm.group(4).lower()] = ((pm.group(6) or "variant").lower(), bool(pm.group(5)))
            continue
        dm = re.match(r"^(?:Dim|Static)\s+(.*)$", code, re.I)
        if dm:
            for part in re.split(r",(?![^()]*\))", dm.group(1)):
                vm = re.match(r"\s*(\w+)(\s*\([^)]*\))?(?:\s+As\s+(?:New\s+)?(\w+))?", part, re.I)
                if vm:
                    types[vm.group(1).lower()] = ((vm.group(3) or "variant").lower(), bool(vm.group(2)))
            continue
        # calls: Name(args) or statement form "Name arg, arg"
        for cname, cargs in paren_calls(code):
            check_call(path, no, cname, cargs, sigs, types, problems)
        pm2 = re.match(r"^(\w+)\s*\((.*)\)\s*$", code)
        if pm2 and pm2.group(1).lower() in sigs and len(split_args(pm2.group(2))) > 1 \
                and code.count("(") == code.count(")") and next(paren_calls(code))[1] == pm2.group(2):
            problems.append(f"{path}:{no}: '{pm2.group(1)}(a, b)' as a statement needs 'Call' or no parentheses - 'Expected: ='")
        sm = re.match(r"^(\w+)\s+(?!=)([^=].*)$", code)
        if sm and sm.group(1).lower() in sigs and not re.match(r"^(If|ElseIf|For|Do|While|Select|Case|With|Set|Dim|ReDim|Call|Exit|End|Next|Loop|GoTo|On|Debug|Option)\b", code, re.I):
            check_call(path, no, sm.group(1), sm.group(2), sigs, types, problems)
    return problems


def split_args(text):
    """Split an argument list at top-level commas (depth-aware)."""
    args, depth, cur = [], 0, ""
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            args.append(cur.strip())
            cur = ""
        else:
            cur += ch
    if cur.strip() or args:
        args.append(cur.strip())
    return args


def paren_calls(code):
    """(name, argument text) for every identifier directly followed by '(' ... ')'."""
    for m in re.finditer(r"(?<![\w.])(\w+)\s*\(", code):
        depth, i = 1, m.end()
        while i < len(code) and depth:
            if code[i] == "(":
                depth += 1
            elif code[i] == ")":
                depth -= 1
            i += 1
        yield m.group(1), code[m.end():i - 1]


def check_call(path, no, name, argtext, sigs, types, problems):
    sig = sigs.get(name.lower())
    if sig is None or name.lower() in types:
        return
    args = split_args(argtext) if argtext.strip() else []
    required = sum(1 for p in sig if p[3] == "required")
    has_paramarray = any(p[3] == "paramarray" for p in sig)
    if len(args) < required or (len(args) > len(sig) and not has_paramarray):
        problems.append(f"{path}:{no}: {name} called with {len(args)} argument(s), expects "
                        f"{required}{'..' + str(len(sig)) if len(sig) != required else ''} - 'Wrong number of arguments'")
    for i, a in enumerate(args):
        if i >= len(sig) or not re.fullmatch(r"\w+", a):
            continue
        mode, ptype, parr, _ = sig[i]
        var = types.get(a.lower())
        if mode != "byref" or var is None:
            continue
        vtype, varr = var
        if parr != varr:
            problems.append(f"{path}:{no}: '{a}' passed to {name} parameter {i + 1}: array/non-array mismatch (ByRef)")
        elif ptype != "variant" and vtype != ptype:
            problems.append(f"{path}:{no}: '{a}' ({vtype}) passed ByRef to {name} parameter {i + 1} ({ptype}) - compile error 'ByRef argument type mismatch'")


def main(argv):
    if len(argv) > 1:
        files = [Path(a) for a in argv[1:]]
    else:
        files = sorted([*(ROOT / "src").rglob("*.bas"), *(ROOT / "tests/vba").glob("*.bas"), *(ROOT / "tools").glob("*.bas")])
    infos = {f: module_info(f) for f in files}
    src_files = [f for f in files if "src" in f.parts]
    problems = []

    # public names must be unique across src modules (tests may reuse helper names privately)
    owner = {}
    for f in src_files:
        _, _, public, _, procs = infos[f]
        dupes = {p for p in procs if procs.count(p) > 1}
        for d in dupes:
            problems.append(f"{rel(f)}: procedure '{d}' defined twice")
        for name in public:
            if name in owner:
                problems.append(f"public '{name}' defined in {owner[name].name} and {f.name} (ambiguous name)")
            owner[name] = f
    module_names = {f.stem.lower() for f in files}
    sigs = proc_signatures([infos[f][1] for f in files])
    for f in files:
        text, lines, public, private, _ = infos[f]
        known = set(owner) | module_names
        if "tests" in f.parts:
            # test modules see the src publics and their own; the runner shim adds PTDict
            known |= public
        problems += check_blocks(rel(f), lines)
        problems += check_physical(rel(f), text)
        problems += check_reserved_names(rel(f), lines)
        problems += check_byref_types(rel(f), lines, sigs)
        problems += check_identifiers(rel(f), lines, known | public, private)
    for p in problems:
        print(p)
    print(f"vba_lint: {len(files)} files, {len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
