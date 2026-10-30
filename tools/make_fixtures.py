#!/usr/bin/env python3
"""Generate the renumbering test documents in tests/fixtures/.

For every fixture <name>.docx this also writes <name>.expected.txt: the visible
paragraph texts after a correct renumbering (TAB kept as TAB) plus the number of
review comments expected ("#flags N" in the first line). The expectation is
computed from the fixture model here, independently of the VBA code.

Requires python-docx.  Usage:  python3 tools/make_fixtures.py
"""
import copy
import random
from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures"
INDENT = Cm(1.25)  # settings.json format.subFeatureIndentCm


# ---------------------------------------------------------------- model
#
# A fixture is a list of items:
#   ("text", str)                         paragraph outside the claims
#   ("claim", label, [parts...])          claim; parts are body paragraphs
# Body paragraphs are lists of segments:
#   str                                   literal text
#   ("ref", i)                            reference to claim index i (0-based, document order)
#   ("range", i, j, word)                 "<i> word <j>"
#   ("del", str) / ("ins", str)           tracked deletion / insertion (tracked fixture only)
#   ("insref", i)                         tracked insertion holding a reference to claim i
# A body paragraph given as ("indent", segments) is an indented sub-feature.


def is_int(s):
    return s.isdigit() and len(s) <= 6


def range_is_safe(old_from, old_to, a, b, index):
    """Mirror of the documented rule (see docs/project-context.md section 4)."""
    if a > b:
        return False
    if is_int(old_from) and is_int(old_to):
        x, y = int(old_from), int(old_to)
        if y - x != b - a:
            return False
        for m in range(x + 1, y):
            k = index.get(str(m))
            if k is not None and not (a <= k <= b):
                return False
    return True


def render(items):
    """Returns (old_paragraphs, expected_paragraphs, expected_flags)."""
    labels = [it[1] for it in items if it[0] == "claim"]
    index = {lab: n + 1 for n, lab in enumerate(labels)}
    old, new, flags = [], [], 0

    def seg_text(segs, renumbered):
        nonlocal flags
        out = []
        for s in segs:
            if isinstance(s, str):
                out.append(s)
            elif s[0] == "ref":
                out.append(str(s[1] + 1) if renumbered else labels[s[1]])
            elif s[0] == "range":
                _, i, j, word = s
                if renumbered:
                    out.append(f"{i + 1}{word}{j + 1}")
                    if not range_is_safe(labels[i], labels[j], i + 1, j + 1, index):
                        flags += 1
                else:
                    out.append(f"{labels[i]}{word}{labels[j]}")
            elif s[0] == "ins":
                out.append(s[1])
            elif s[0] == "insref":
                out.append(str(s[1] + 1) if renumbered else labels[s[1]])
            elif s[0] == "del":
                pass  # deleted text is not visible
        return "".join(out)

    claim_no = 0
    for it in items:
        if it[0] == "text":
            old.append(it[1])
            new.append(it[1])
            continue
        claim_no += 1
        _, label, parts = it
        for k, part in enumerate(parts):
            segs = part[1] if isinstance(part, tuple) and part[0] == "indent" else part
            o, n = seg_text(segs, False), seg_text(segs, True)
            if k == 0:
                o, n = f"{label}.\t{o}", f"{claim_no}.\t{n}"
            old.append(o)
            new.append(n)
        old.append("")
        new.append("")
    if old and old[-1] == "":
        old.pop()
        new.pop()
    return old, new, flags


# ---------------------------------------------------------------- docx writing

def tracked(tag, text, rid):
    el = OxmlElement(f"w:{tag}")
    el.set(qn("w:id"), str(rid))
    el.set(qn("w:author"), "Fixture")
    el.set(qn("w:date"), "2026-10-30T00:00:00Z")
    r = OxmlElement("w:r")
    t = OxmlElement("w:delText" if tag == "del" else "w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    el.append(r)
    return el


def write_docx(items, path):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    labels = [it[1] for it in items if it[0] == "claim"]
    rid = [1]

    def add_par(segs, prefix=None, indent=False):
        p = doc.add_paragraph()
        if indent:
            p.paragraph_format.left_indent = INDENT
        if prefix is not None:
            p.add_run(prefix)
        for s in segs:
            if isinstance(s, str):
                p.add_run(s)
            elif s[0] == "ref":
                p.add_run(labels[s[1]])
            elif s[0] == "range":
                p.add_run(f"{labels[s[1]]}{s[3]}{labels[s[2]]}")
            elif s[0] == "insref":
                p._p.append(tracked("ins", labels[s[1]], rid[0]))
                rid[0] += 1
            else:
                p._p.append(tracked(s[0], s[1], rid[0]))
                rid[0] += 1
        return p

    for it in items:
        if it[0] == "text":
            doc.add_paragraph(it[1])
            continue
        _, label, parts = it
        for k, part in enumerate(parts):
            indent = isinstance(part, tuple) and part[0] == "indent"
            segs = part[1] if indent else part
            add_par(segs, f"{label}.\t" if k == 0 else None, indent)
        doc.add_paragraph("")
    body = doc.element.body
    last = body.findall(qn("w:p"))[-1]
    if not last.findall(".//" + qn("w:t")):
        body.remove(last)
    doc.save(path)


# ---------------------------------------------------------------- fixtures

def fixture_en_small():
    return [
        ("text", "Claims"),
        ("claim", "1", [["A method for processing a signal, the method comprising:"],
                        ("indent", ["receiving a signal;"]),
                        ("indent", ["filtering the signal; and"]),
                        ("indent", ["outputting the filtered signal."])]),
        ("claim", "1a", [["The method of claim ", ("ref", 0), ", wherein the filtering is a low-pass filtering."]]),
        ("claim", "1.1", [["The method of claim ", ("ref", 1), ", wherein a cut-off frequency is adjustable."]]),
        ("claim", "2", [["The method of any one of claims ", ("range", 0, 2, " to "), ", wherein the signal is digital."]]),
        ("claim", "1a1", [["The method of claim ", ("ref", 5), " or ", ("ref", 3), ", further comprising:"],
                          ("indent", ["storing the signal according to claim ", ("ref", 2), "."])]),
        ("claim", "1a.1", [["The method of claim ", ("ref", 4), ", wherein the storing is volatile."]]),
        ("claim", "3", [["The method of claims ", ("range", 0, 3, "-"), " and/or ", ("ref", 5), "."]]),
        ("claim", "4", [["A device configured to carry out the method of claim ", ("ref", 0), "."]]),
        ("claim", "4a", [["The device of claim ", ("ref", 7), ", comprising a processor, wherein the processor is configured as claimed in claim 12."]]),
    ]


def fixture_de_small():
    return [
        ("text", "Patentansprüche"),
        ("claim", "1", [["Verfahren zur Verarbeitung eines Signals, mit:"],
                        ("indent", ["Empfangen eines Signals;"]),
                        ("indent", ["Filtern des Signals; und"]),
                        ("indent", ["Ausgeben des gefilterten Signals."])]),
        ("claim", "1a", [["Verfahren nach Anspruch ", ("ref", 0), ", wobei das Filtern ein Tiefpassfiltern ist."]]),
        ("claim", "2", [["Verfahren nach einem der Ansprüche ", ("range", 0, 1, " bis "), ", wobei das Signal digital ist."]]),
        ("claim", "2.1", [["Verfahren nach Anspruch ", ("ref", 2), " oder ", ("ref", 1), ", ferner mit:"],
                          ("indent", ["Speichern des Signals gemäß Patentanspruch ", ("ref", 1), "."])]),
        ("claim", "3", [["Vorrichtung, eingerichtet zur Durchführung des Verfahrens nach einem der Ansprüche ",
                         ("range", 0, 2, " bis "), "."]]),
        ("claim", "3a", [["Vorrichtung nach Anspruch ", ("ref", 4), ", mit einem Prozessor gemäß den Ansprüchen ",
                          ("ref", 0), ", ", ("ref", 2), " und ", ("ref", 3), " sowie des Anspruchs ", ("ref", 4), "."]]),
        ("claim", "4", [["Computerprogramm mit Befehlen, die bewirken, dass die Vorrichtung des Anspruchs ",
                         ("ref", 5), " das Verfahren nach Anspruch ", ("ref", 0), " ausführt."]]),
    ]


def fixture_tracked():
    return [
        ("text", "Claims"),
        ("claim", "1", [["A method comprising a ", ("del", "first "), ("ins", "single "), "step."]]),
        ("claim", "1a", [["The method of claim ", ("ref", 0), ", wherein the step is repeated."]]),
        ("claim", "2", [["The method of claim ", ("del", "1"), ("insref", 1), ", wherein the step is timed."]]),
        ("claim", "3", [["The method of claim ", ("ref", 2), ("del", " or claim 1"), "."]]),
    ]


def fixture_large(n_main=160, seed=20187485):
    """~200 claims: integers 1..n_main plus supplementary claims at random places."""
    rnd = random.Random(seed)
    labels = []
    for k in range(1, n_main + 1):
        labels.append(str(k))
        r = rnd.random()
        if r < 0.12:
            labels.append(f"{k}a")
        elif r < 0.20:
            labels += [f"{k}.1", f"{k}a.1"]
        elif r < 0.24:
            labels += [f"{k}a1", f"{k}b"]
    items = [("text", "Claims")]
    for idx, lab in enumerate(labels):
        if idx == 0 or rnd.random() < 0.06:
            parts = [["A method for operating a node, the method comprising:"],
                     ("indent", ["determining a state;"]),
                     ("indent", ["transmitting the state."])]
        else:
            t = rnd.randrange(0, idx)
            choice = rnd.random()
            if choice < 0.55:
                body = ["The method of claim ", ("ref", t), ", wherein the state is cached."]
            elif choice < 0.75:
                u = rnd.randrange(0, idx)
                body = ["The method of claim ", ("ref", min(t, u)), " or ", ("ref", max(t, u)), ", further comprising a timer."]
            elif choice < 0.9:
                u = rnd.randrange(t, idx)
                body = ["The method of any one of claims ", ("range", t, u, " to "), ", wherein the node is mobile."]
            else:
                body = ["The method of claim ", ("ref", t), ", further comprising:"]
                parts = [body, ("indent", ["repeating the determining according to claim ", ("ref", rnd.randrange(0, idx)), "."])]
                items.append(("claim", lab, parts))
                continue
            parts = [body]
        items.append(("claim", lab, parts))
    return items


FIXTURES = {
    # extra flags not visible in the model: EN small refers to a non-existent "claim 12"
    "renumber_EN_small": (fixture_en_small, 1),
    "renumber_DE_small": (fixture_de_small, 0),
    "renumber_tracked_changes": (fixture_tracked, 0),
    "renumber_EN_large": (fixture_large, 0),
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, (make, extra_flags) in FIXTURES.items():
        items = make()
        old, new, flags = render(items)
        write_docx(items, OUT / f"{name}.docx")
        flags += extra_flags
        (OUT / f"{name}.expected.txt").write_text(
            f"#flags {flags}\n" + "\n".join(new) + "\n", encoding="utf-8")
        n_claims = sum(1 for it in items if it[0] == "claim")
        print(f"{name}: {n_claims} claims, {len(old)} paragraphs, {flags} expected comment(s)")


if __name__ == "__main__":
    main()
