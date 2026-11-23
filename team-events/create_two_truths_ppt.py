#!/usr/bin/env python3
"""Generate Two Truths and a Lie game PPT + separate solutions PPT."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

C = {
    "bg": "F7F9F8",
    "ink": "1A2B2E",
    "muted": "4A5C60",
    "teal": "0D7377",
    "teal_dark": "095456",
    "coral": "E85A4F",
    "sand": "E8F0EF",
    "white": "FFFFFF",
    "line": "C5D5D3",
    "green": "1B7A4E",
}

# Fictional warm-up rounds (solvable). lie_index is 1-based.
ROUNDS = [
    {
        "name": "Alex — Product",
        "statements": [
            "I once slept overnight in an airport.",
            "I can solve a Rubik's cube in under 2 minutes.",
            "I have never broken a bone.",
        ],
        "lie": 2,
        "why": "Alex cannot solve a Rubik's cube — that statement is the lie.",
    },
    {
        "name": "Jordan — Engineering",
        "statements": [
            "I speak three languages.",
            "I hate chocolate.",
            "I have visited Japan.",
        ],
        "lie": 2,
        "why": "Jordan loves chocolate — 'I hate chocolate' is the lie.",
    },
    {
        "name": "Sam — Design",
        "statements": [
            "I have a pet named after a cartoon character.",
            "I was on local TV as a kid.",
            "I have never flown in a plane.",
        ],
        "lie": 3,
        "why": "Sam has flown many times — 'never flown' is the lie.",
    },
    {
        "name": "Riley — Data",
        "statements": [
            "I ran a half-marathon last year.",
            "I can play the guitar.",
            "I was born on February 29 (leap day).",
        ],
        "lie": 3,
        "why": "Riley's birthday is not on a leap day — that is the lie.",
    },
    {
        "name": "Casey — Marketing",
        "statements": [
            "I once met a celebrity in an elevator.",
            "I collect vintage stamps.",
            "I have skydived.",
        ],
        "lie": 2,
        "why": "Casey does not collect stamps — that is the lie.",
    },
    {
        "name": "Morgan — Ops",
        "statements": [
            "I am left-handed.",
            "I lived on a boat for a month.",
            "I have never watched a Marvel movie.",
        ],
        "lie": 3,
        "why": "Morgan has watched Marvel movies — that claim is the lie.",
    },
    {
        "name": "Taylor — Sales",
        "statements": [
            "I have broken a bone twice.",
            "I can whistle really loudly.",
            "I once worked as a barista.",
        ],
        "lie": 1,
        "why": "Taylor has never broken a bone — statement 1 is the lie.",
    },
    {
        "name": "Jamie — Support",
        "statements": [
            "I have a twin sibling.",
            "I prefer tea over coffee.",
            "I know a magic card trick.",
        ],
        "lie": 1,
        "why": "Jamie does not have a twin — that is the lie.",
    },
]


def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb(color)
    bg.line.fill.background()
    sp = bg._element
    tree = slide.shapes._spTree
    tree.remove(sp)
    tree.insert(2, sp)


def rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    return s


def round_rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    try:
        s.adjustments[0] = 0.08
    except Exception:
        pass
    return s


def run_style(run, size=18, bold=False, color=None, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    if color:
        run.font.color.rgb = rgb(color)


def textbox(slide, l, t, w, h, text, size=18, bold=False, color=None, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    run_style(r, size=size, bold=bold, color=color or C["ink"], font=font)
    return box


def bullets(slide, l, t, w, h, items, size=15, color=None, spacing=8, numbered=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = (f"{i + 1}.  {item}" if numbered else f"•  {item}")
        run_style(r, size=size, color=color or C["ink"])
    return box


def accent(slide):
    rect(slide, 0, 0, Inches(13.333), Inches(0.12), C["teal"])


def footer(slide, page, total, label="Two Truths and a Lie"):
    textbox(slide, Inches(0.5), Inches(7.1), Inches(10), Inches(0.3), label, size=11, color=C["muted"])
    textbox(slide, Inches(11.5), Inches(7.1), Inches(1.3), Inches(0.3), f"{page} / {total}", size=11, color=C["muted"], align=PP_ALIGN.RIGHT)


def new_slide(prs, blank, bg=None):
    s = prs.slides.add_slide(blank)
    set_bg(s, bg or C["bg"])
    accent(s)
    return s


def build_game():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    # title + how + scoring + 8 rounds + live play + closing = 13
    total = 13

    # Title
    s = new_slide(prs, blank, C["teal_dark"])
    rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    rect(s, 0, Inches(5.9), Inches(13.333), Inches(1.6), C["teal"])
    textbox(s, Inches(0.8), Inches(2.1), Inches(11), Inches(0.4), "TEAM EVENT GAME", size=15, bold=True, color="A8D5D7")
    textbox(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(1.2), "Two Truths and a Lie", size=44, bold=True, color=C["white"], font="Georgia")
    textbox(s, Inches(0.8), Inches(4.0), Inches(10), Inches(0.6), "Share 3 statements. The group finds the lie.", size=20, color="D0E8E9")
    textbox(s, Inches(0.8), Inches(6.25), Inches(11), Inches(0.4), "Player deck  ·  Solutions are in a separate file — do not project those", size=14, color=C["white"])

    # How to play
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "HOW TO PLAY", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Rules in 60 seconds", size=32, bold=True, color=C["ink"], font="Georgia")
    steps = [
        "Each speaker prepares 2 true facts and 1 lie about themselves (work-safe).",
        "They read all three statements — same tone, no hints.",
        "The group discusses briefly, then votes which one is the lie.",
        "Reveal the lie. Correct guessers score 1 point.",
        "Speaker scores 1 point if most of the group guessed wrong.",
        "Next person goes. Keep rounds short and fun.",
    ]
    for i, step in enumerate(steps):
        y = Inches(1.45 + i * 0.85)
        round_rect(s, Inches(0.55), y, Inches(12.2), Inches(0.72), C["white"], C["line"])
        rect(s, Inches(0.55), y, Inches(0.15), Inches(0.72), C["coral"] if i % 2 == 0 else C["teal"])
        textbox(s, Inches(0.95), y + Inches(0.18), Inches(0.6), Inches(0.4), f"{i + 1}", size=18, bold=True, color=C["teal"])
        textbox(s, Inches(1.6), y + Inches(0.18), Inches(10.8), Inches(0.45), step, size=16, color=C["ink"])
    footer(s, 2, total)

    # Scoring + tips
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "SCORING & TIPS", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Keep it fair and light", size=32, bold=True, color=C["ink"], font="Georgia")
    left = [
        "Correct guess = 1 point",
        "Fool the majority = speaker gets 1 point",
        "Optional team mode: score as groups",
        "Timebox: ~2 minutes per person",
    ]
    right = [
        "Keep statements work-safe and kind",
        "Anyone may pass — no pressure",
        "Remote: use Zoom polls or chat votes",
        "Host models first with a sample round",
    ]
    round_rect(s, Inches(0.55), Inches(1.5), Inches(5.9), Inches(5.0), C["white"], C["line"])
    textbox(s, Inches(0.85), Inches(1.75), Inches(5.3), Inches(0.4), "Scoring", size=20, bold=True, color=C["coral"])
    bullets(s, Inches(0.85), Inches(2.4), Inches(5.3), Inches(3.8), left, size=16, spacing=14)
    round_rect(s, Inches(6.85), Inches(1.5), Inches(5.9), Inches(5.0), C["sand"])
    textbox(s, Inches(7.15), Inches(1.75), Inches(5.3), Inches(0.4), "Facilitator tips", size=20, bold=True, color=C["teal"])
    bullets(s, Inches(7.15), Inches(2.4), Inches(5.3), Inches(3.8), right, size=16, spacing=14)
    footer(s, 3, total)

    # Warm-up intro
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "WARM-UP ROUNDS", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(2.2), Inches(12), Inches(0.8), "Practice with fictional teammates", size=36, bold=True, color=C["ink"], font="Georgia")
    textbox(
        s,
        Inches(0.6),
        Inches(3.2),
        Inches(11),
        Inches(1.2),
        "Vote on which statement is the lie for each character.\nAnswers are in the separate Solutions deck — keep that closed for now.",
        size=18,
        color=C["muted"],
    )
    footer(s, 4, total)

    # One slide per round
    for i, rnd in enumerate(ROUNDS):
        s = new_slide(prs, blank)
        textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), f"ROUND {i + 1} OF {len(ROUNDS)}", size=12, bold=True, color=C["teal"])
        textbox(s, Inches(0.6), Inches(0.75), Inches(12), Inches(0.55), rnd["name"], size=30, bold=True, color=C["ink"], font="Georgia")
        textbox(s, Inches(0.6), Inches(1.35), Inches(12), Inches(0.35), "Which statement is the LIE? Vote 1, 2, or 3.", size=16, color=C["muted"])
        for j, stmt in enumerate(rnd["statements"]):
            y = Inches(1.95 + j * 1.5)
            round_rect(s, Inches(0.55), y, Inches(12.2), Inches(1.25), C["white"], C["line"])
            rect(s, Inches(0.55), y, Inches(1.1), Inches(1.25), C["teal"] if j != 1 else C["coral"])
            textbox(s, Inches(0.55), y + Inches(0.35), Inches(1.1), Inches(0.55), str(j + 1), size=28, bold=True, color=C["white"], align=PP_ALIGN.CENTER)
            textbox(s, Inches(1.95), y + Inches(0.38), Inches(10.4), Inches(0.55), stmt, size=22, color=C["ink"])
        footer(s, 5 + i, total)

    # Live play
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "YOUR TURN", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.75), Inches(12), Inches(0.55), "Now play with the real team", size=32, bold=True, color=C["ink"], font="Georgia")
    prompts = [
        "Write your 3 statements privately (2 truths + 1 lie).",
        "Volunteer order: raise hands or random pick.",
        "Speaker reads → group votes → reveal → score.",
        "Aim for 8–12 people if time is ~20 minutes.",
    ]
    for i, p in enumerate(prompts):
        col, row = i % 2, i // 2
        x = Inches(0.55 + col * 6.35)
        y = Inches(1.7 + row * 2.3)
        round_rect(s, x, y, Inches(6.05), Inches(2.0), C["sand"] if row == 0 else C["white"], C["line"])
        textbox(s, x + Inches(0.3), y + Inches(0.35), Inches(1), Inches(0.4), f"0{i + 1}", size=22, bold=True, color=C["coral"])
        textbox(s, x + Inches(0.3), y + Inches(0.95), Inches(5.4), Inches(0.8), p, size=16, color=C["ink"])
    footer(s, 13, total)  # will fix - actually rounds are 5-12, live is 13? Let's recount
    # title=1, how=2, scoring=3, warmup intro=4, rounds 5-12 (8), live=13. Need closing too.
    # I'll add closing as 14 and fix totals.

    # Closing
    s = new_slide(prs, blank, C["teal_dark"])
    rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    textbox(s, Inches(0.8), Inches(2.5), Inches(11.5), Inches(1.0), "Ready to find the lies?", size=40, bold=True, color=C["white"], font="Georgia")
    textbox(s, Inches(0.8), Inches(3.7), Inches(11), Inches(0.8), "Facilitator: open the Solutions deck only after voting.\nDo not share that file with the room.", size=18, color="D0E8E9")

    # Fix footers: recount actual slides
    n = len(prs.slides)
    # Re-stamp footers is hard; just save with consistent TOTAL in later pass
    # Actually live slide said 13/13 but we have closing = 14. Let me rebuild footer numbers more carefully by regenerating.

    out = "/agent/repos/course/team-events/Two_Truths_and_a_Lie.pptx"
    art = "/opt/cursor/artifacts/Two_Truths_and_a_Lie.pptx"
    prs.save(out)
    prs.save(art)
    print(f"Game deck: {out} ({n} slides)")
    return out, n


def build_solutions():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total = 2 + len(ROUNDS)  # title + overview + one per round? Or compact: title + all answers + notes
    # Compact: title, all answers on 2 slides, how to use = 4 slides
    total = 4

    s = new_slide(prs, blank, C["teal_dark"])
    rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    textbox(s, Inches(0.8), Inches(2.0), Inches(11), Inches(0.4), "FACILITATOR ONLY — DO NOT PROJECT", size=15, bold=True, color=C["coral"])
    textbox(s, Inches(0.8), Inches(2.7), Inches(11.5), Inches(1.2), "Two Truths and a Lie\nSolutions", size=40, bold=True, color=C["white"], font="Georgia")
    textbox(s, Inches(0.8), Inches(5.0), Inches(11), Inches(0.6), "Keep this deck closed until after the group votes each round.", size=16, color="D0E8E9")

    # All answers overview
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "ANSWER KEY", size=12, bold=True, color=C["coral"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.5), "Warm-up rounds — which is the lie?", size=28, bold=True, color=C["ink"], font="Georgia")
    for i, rnd in enumerate(ROUNDS[:4]):
        y = Inches(1.4 + i * 1.3)
        round_rect(s, Inches(0.55), y, Inches(12.2), Inches(1.15), C["white"], C["line"])
        textbox(s, Inches(0.8), y + Inches(0.15), Inches(8), Inches(0.35), f"Round {i + 1}: {rnd['name']}", size=15, bold=True, color=C["ink"])
        textbox(s, Inches(0.8), y + Inches(0.55), Inches(11.5), Inches(0.45), f"LIE = Statement {rnd['lie']}  —  {rnd['statements'][rnd['lie'] - 1]}", size=14, color=C["coral"])
    footer(s, 2, total, "Solutions — Facilitator only")

    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "ANSWER KEY (CONTINUED)", size=12, bold=True, color=C["coral"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.5), "Warm-up rounds 5–8", size=28, bold=True, color=C["ink"], font="Georgia")
    for i, rnd in enumerate(ROUNDS[4:]):
        idx = i + 5
        y = Inches(1.4 + i * 1.3)
        round_rect(s, Inches(0.55), y, Inches(12.2), Inches(1.15), C["white"], C["line"])
        textbox(s, Inches(0.8), y + Inches(0.15), Inches(8), Inches(0.35), f"Round {idx}: {rnd['name']}", size=15, bold=True, color=C["ink"])
        textbox(s, Inches(0.8), y + Inches(0.55), Inches(11.5), Inches(0.45), f"LIE = Statement {rnd['lie']}  —  {rnd['statements'][rnd['lie'] - 1]}", size=14, color=C["coral"])
    footer(s, 3, total, "Solutions — Facilitator only")

    # Detailed why
    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "REVEAL NOTES", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.5), "What to say when revealing", size=28, bold=True, color=C["ink"], font="Georgia")
    items = [f"R{i + 1}: {r['why']}" for i, r in enumerate(ROUNDS)]
    # two columns
    mid = 4
    bullets(s, Inches(0.6), Inches(1.4), Inches(6.0), Inches(5.2), items[:mid], size=14, spacing=12)
    bullets(s, Inches(6.9), Inches(1.4), Inches(6.0), Inches(5.2), items[mid:], size=14, spacing=12)
    footer(s, 4, total, "Solutions — Facilitator only")

    out = "/agent/repos/course/team-events/Two_Truths_and_a_Lie_SOLUTIONS.pptx"
    art = "/opt/cursor/artifacts/Two_Truths_and_a_Lie_SOLUTIONS.pptx"
    prs.save(out)
    prs.save(art)
    print(f"Solutions deck: {out} ({len(prs.slides)} slides)")
    return out


def rebuild_game_with_correct_footers():
    """Rebuild game deck with accurate page numbers."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total = 14

    s = new_slide(prs, blank, C["teal_dark"])
    rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    rect(s, 0, Inches(5.9), Inches(13.333), Inches(1.6), C["teal"])
    textbox(s, Inches(0.8), Inches(2.1), Inches(11), Inches(0.4), "TEAM EVENT GAME", size=15, bold=True, color="A8D5D7")
    textbox(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(1.2), "Two Truths and a Lie", size=44, bold=True, color=C["white"], font="Georgia")
    textbox(s, Inches(0.8), Inches(4.0), Inches(10), Inches(0.6), "Share 3 statements. The group finds the lie.", size=20, color="D0E8E9")
    textbox(s, Inches(0.8), Inches(6.25), Inches(11), Inches(0.4), "Player deck  ·  Solutions are in a separate file — do not project those", size=14, color=C["white"])

    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "HOW TO PLAY", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Rules in 60 seconds", size=32, bold=True, color=C["ink"], font="Georgia")
    steps = [
        "Each speaker prepares 2 true facts and 1 lie about themselves (work-safe).",
        "They read all three statements — same tone, no hints.",
        "The group discusses briefly, then votes which one is the lie.",
        "Reveal the lie. Correct guessers score 1 point.",
        "Speaker scores 1 point if most of the group guessed wrong.",
        "Next person goes. Keep rounds short and fun.",
    ]
    for i, step in enumerate(steps):
        y = Inches(1.45 + i * 0.85)
        round_rect(s, Inches(0.55), y, Inches(12.2), Inches(0.72), C["white"], C["line"])
        rect(s, Inches(0.55), y, Inches(0.15), Inches(0.72), C["coral"] if i % 2 == 0 else C["teal"])
        textbox(s, Inches(0.95), y + Inches(0.18), Inches(0.6), Inches(0.4), f"{i + 1}", size=18, bold=True, color=C["teal"])
        textbox(s, Inches(1.6), y + Inches(0.18), Inches(10.8), Inches(0.45), step, size=16, color=C["ink"])
    footer(s, 2, total)

    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "SCORING & TIPS", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Keep it fair and light", size=32, bold=True, color=C["ink"], font="Georgia")
    left = [
        "Correct guess = 1 point",
        "Fool the majority = speaker gets 1 point",
        "Optional team mode: score as groups",
        "Timebox: ~2 minutes per person",
    ]
    right = [
        "Keep statements work-safe and kind",
        "Anyone may pass — no pressure",
        "Remote: use Zoom polls or chat votes",
        "Host models first with a sample round",
    ]
    round_rect(s, Inches(0.55), Inches(1.5), Inches(5.9), Inches(5.0), C["white"], C["line"])
    textbox(s, Inches(0.85), Inches(1.75), Inches(5.3), Inches(0.4), "Scoring", size=20, bold=True, color=C["coral"])
    bullets(s, Inches(0.85), Inches(2.4), Inches(5.3), Inches(3.8), left, size=16, spacing=14)
    round_rect(s, Inches(6.85), Inches(1.5), Inches(5.9), Inches(5.0), C["sand"])
    textbox(s, Inches(7.15), Inches(1.75), Inches(5.3), Inches(0.4), "Facilitator tips", size=20, bold=True, color=C["teal"])
    bullets(s, Inches(7.15), Inches(2.4), Inches(5.3), Inches(3.8), right, size=16, spacing=14)
    footer(s, 3, total)

    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "WARM-UP ROUNDS", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(2.2), Inches(12), Inches(0.8), "Practice with fictional teammates", size=36, bold=True, color=C["ink"], font="Georgia")
    textbox(
        s,
        Inches(0.6),
        Inches(3.2),
        Inches(11),
        Inches(1.2),
        "Vote on which statement is the lie for each character.\nAnswers are in the separate Solutions deck — keep that closed for now.",
        size=18,
        color=C["muted"],
    )
    footer(s, 4, total)

    for i, rnd in enumerate(ROUNDS):
        s = new_slide(prs, blank)
        textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), f"ROUND {i + 1} OF {len(ROUNDS)}", size=12, bold=True, color=C["teal"])
        textbox(s, Inches(0.6), Inches(0.75), Inches(12), Inches(0.55), rnd["name"], size=30, bold=True, color=C["ink"], font="Georgia")
        textbox(s, Inches(0.6), Inches(1.35), Inches(12), Inches(0.35), "Which statement is the LIE? Vote 1, 2, or 3.", size=16, color=C["muted"])
        for j, stmt in enumerate(rnd["statements"]):
            y = Inches(1.95 + j * 1.5)
            fill_bar = C["coral"] if j == 1 else C["teal"]
            round_rect(s, Inches(0.55), y, Inches(12.2), Inches(1.25), C["white"], C["line"])
            rect(s, Inches(0.55), y, Inches(1.1), Inches(1.25), fill_bar)
            textbox(s, Inches(0.55), y + Inches(0.35), Inches(1.1), Inches(0.55), str(j + 1), size=28, bold=True, color=C["white"], align=PP_ALIGN.CENTER)
            textbox(s, Inches(1.95), y + Inches(0.38), Inches(10.4), Inches(0.55), stmt, size=22, color=C["ink"])
        footer(s, 5 + i, total)

    s = new_slide(prs, blank)
    textbox(s, Inches(0.6), Inches(0.35), Inches(10), Inches(0.3), "YOUR TURN", size=12, bold=True, color=C["teal"])
    textbox(s, Inches(0.6), Inches(0.75), Inches(12), Inches(0.55), "Now play with the real team", size=32, bold=True, color=C["ink"], font="Georgia")
    prompts = [
        "Write your 3 statements privately (2 truths + 1 lie).",
        "Volunteer order: raise hands or random pick.",
        "Speaker reads → group votes → reveal → score.",
        "Aim for 8–12 people if time is ~20 minutes.",
    ]
    for i, p in enumerate(prompts):
        col, row = i % 2, i // 2
        x = Inches(0.55 + col * 6.35)
        y = Inches(1.7 + row * 2.3)
        round_rect(s, x, y, Inches(6.05), Inches(2.0), C["sand"] if row == 0 else C["white"], C["line"])
        textbox(s, x + Inches(0.3), y + Inches(0.35), Inches(1), Inches(0.4), f"0{i + 1}", size=22, bold=True, color=C["coral"])
        textbox(s, x + Inches(0.3), y + Inches(0.95), Inches(5.4), Inches(0.8), p, size=16, color=C["ink"])
    footer(s, 13, total)

    s = new_slide(prs, blank, C["teal_dark"])
    rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    textbox(s, Inches(0.8), Inches(2.5), Inches(11.5), Inches(1.0), "Ready to find the lies?", size=40, bold=True, color=C["white"], font="Georgia")
    textbox(s, Inches(0.8), Inches(3.7), Inches(11), Inches(0.8), "Facilitator: open the Solutions deck only after voting.\nDo not share that file with the room.", size=18, color="D0E8E9")

    out = "/agent/repos/course/team-events/Two_Truths_and_a_Lie.pptx"
    art = "/opt/cursor/artifacts/Two_Truths_and_a_Lie.pptx"
    prs.save(out)
    prs.save(art)
    print(f"Game deck: {out} ({len(prs.slides)} slides)")
    return out


if __name__ == "__main__":
    rebuild_game_with_correct_footers()
    build_solutions()
    # zip for easy download
    import zipfile
    zpath = "/opt/cursor/artifacts/Two_Truths_and_a_Lie_Pack.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(
            "/agent/repos/course/team-events/Two_Truths_and_a_Lie.pptx",
            "Two_Truths_and_a_Lie.pptx",
        )
        zf.write(
            "/agent/repos/course/team-events/Two_Truths_and_a_Lie_SOLUTIONS.pptx",
            "Two_Truths_and_a_Lie_SOLUTIONS.pptx",
        )
    print(f"Zip: {zpath}")
