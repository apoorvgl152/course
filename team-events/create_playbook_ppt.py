#!/usr/bin/env python3
"""Generate Monthly Team Event Games PowerPoint playbook."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Visual direction: deep teal + coral on soft off-white
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
}

TOTAL = 19


def rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_slide_bg(slide, hex_color: str) -> None:
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb(hex_color)
    bg.line.fill.background()
    sp_tree = slide.shapes._spTree
    sp = bg._element
    sp_tree.remove(sp)
    sp_tree.insert(2, sp)


def add_rect(slide, left, top, width, height, fill_hex, line_hex=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill_hex)
    if line_hex:
        shape.line.color.rgb = rgb(line_hex)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_round_rect(slide, left, top, width, height, fill_hex, line_hex=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill_hex)
    if line_hex:
        shape.line.color.rgb = rgb(line_hex)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def set_run(run, size=18, bold=False, color=None, font_name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font_name
    if color:
        run.font.color.rgb = rgb(color)


def add_textbox(
    slide,
    left,
    top,
    width,
    height,
    text,
    size=18,
    bold=False,
    color=None,
    align=PP_ALIGN.LEFT,
    font_name="Calibri",
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color or C["ink"], font_name=font_name)
    return box


def add_bullets(slide, left, top, width, height, items, size=16, color=None, spacing=8):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = "•  " + item
        set_run(run, size=size, color=color or C["ink"])
    return box


def add_numbered(slide, left, top, width, height, items, size=16, color=None, spacing=10):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = f"{i + 1}.  {item}"
        set_run(run, size=size, color=color or C["ink"])
    return box


def footer(slide, page: int) -> None:
    add_textbox(
        slide,
        Inches(0.5),
        Inches(7.1),
        Inches(10),
        Inches(0.3),
        "Monthly Team Fun Event  ·  Facilitator Deck",
        size=11,
        color=C["muted"],
    )
    add_textbox(
        slide,
        Inches(11.5),
        Inches(7.1),
        Inches(1.3),
        Inches(0.3),
        f"{page} / {TOTAL}",
        size=11,
        color=C["muted"],
        align=PP_ALIGN.RIGHT,
    )


def accent_bar(slide) -> None:
    add_rect(slide, 0, 0, Inches(13.333), Inches(0.12), C["teal"])


def section_label(slide, text: str) -> None:
    add_textbox(
        slide,
        Inches(0.6),
        Inches(0.35),
        Inches(10),
        Inches(0.35),
        text.upper(),
        size=12,
        bold=True,
        color=C["teal"],
    )


def new_slide(prs, blank, bg=None):
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, bg or C["bg"])
    accent_bar(s)
    return s


def build() -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # 1. Title
    s = new_slide(prs, blank, C["teal_dark"])
    add_rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    add_rect(s, 0, Inches(5.8), Inches(13.333), Inches(1.7), C["teal"])
    add_textbox(
        s,
        Inches(0.8),
        Inches(2.0),
        Inches(11),
        Inches(0.5),
        "MONTHLY TEAM EVENT",
        size=16,
        bold=True,
        color="A8D5D7",
    )
    add_textbox(
        s,
        Inches(0.8),
        Inches(2.5),
        Inches(11.5),
        Inches(1.4),
        "Fun Games Playbook",
        size=48,
        bold=True,
        color=C["white"],
        font_name="Georgia",
    )
    add_textbox(
        s,
        Inches(0.8),
        Inches(4.1),
        Inches(10),
        Inches(0.8),
        "Bingo, trivia, scavenger hunts & more — ready-to-run activities\nwith facilitator scripts, rules, and answer keys.",
        size=18,
        color="D0E8E9",
    )
    add_textbox(
        s,
        Inches(0.8),
        Inches(6.2),
        Inches(11),
        Inches(0.5),
        "For remote, hybrid, or in-person teams  ·  30–90 minute sessions",
        size=14,
        color=C["white"],
    )

    # 2. Agenda
    s = new_slide(prs, blank)
    section_label(s, "Overview")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.7),
        Inches(12),
        Inches(0.7),
        "What's inside this deck",
        size=32,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    games_overview = [
        ("01", "Human Bingo", "Mix & mingle icebreaker"),
        ("02", "Two Truths & a Lie", "Storytelling + guessing"),
        ("03", "Team Trivia Sprint", "Quiz rounds with answers"),
        ("04", "Photo Scavenger Hunt", "Creative team challenge"),
        ("05", "Emoji Charades", "Act it out / guess it"),
        ("06", "Minute Challenges", "Fast physical / desk games"),
        ("07", "Would You Rather", "Debate & bond"),
        ("08", "Office Olympics", "Relay-style mini sports"),
    ]
    for i, (num, title, desc) in enumerate(games_overview):
        col, row = i % 4, i // 4
        x = Inches(0.6 + col * 3.15)
        y = Inches(1.7 + row * 2.4)
        add_round_rect(s, x, y, Inches(2.95), Inches(2.1), C["white"], C["line"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.25), Inches(2.5), Inches(0.4), num, size=20, bold=True, color=C["coral"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.75), Inches(2.5), Inches(0.5), title, size=16, bold=True, color=C["ink"])
        add_textbox(s, x + Inches(0.2), y + Inches(1.3), Inches(2.5), Inches(0.5), desc, size=13, color=C["muted"])
    footer(s, 2)

    # 3. How to host
    s = new_slide(prs, blank)
    section_label(s, "Facilitator tips")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.7),
        Inches(12),
        Inches(0.6),
        "How to run a great session",
        size=32,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    tips = [
        ("Pick 2–3 games", "Don't overstuff. One icebreaker + one main game + one closer is plenty for 45–60 min."),
        ("Explain → Demo → Play", "State the goal, show one example, then start. Keep rules under 60 seconds."),
        ("Mix energy levels", "Pair a high-energy game (Bingo/Charades) with a calmer one (Would You Rather)."),
        ("Celebrate loudly", "Cheer winners, shout-outs for creativity, optional small prizes (snacks, stickers, bragging rights)."),
        ("Remote-friendly tweaks", "Use breakout rooms, shared docs, Zoom reactions, and a shared slide for timers."),
        ("Safety & inclusion", "Avoid personal/medical topics, and offer opt-outs without pressure."),
    ]
    for i, (t, d) in enumerate(tips):
        col, row = i % 3, i // 3
        x = Inches(0.55 + col * 4.2)
        y = Inches(1.55 + row * 2.5)
        add_round_rect(s, x, y, Inches(3.95), Inches(2.2), C["sand"])
        add_textbox(s, x + Inches(0.25), y + Inches(0.3), Inches(3.4), Inches(0.5), t, size=16, bold=True, color=C["teal"])
        add_textbox(s, x + Inches(0.25), y + Inches(0.9), Inches(3.4), Inches(1.1), d, size=14, color=C["ink"])
    footer(s, 3)

    # 4. Human Bingo
    s = new_slide(prs, blank)
    section_label(s, "Game 01  ·  15–20 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(10), Inches(0.6), "Human Bingo", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.35),
        Inches(12),
        Inches(0.4),
        "A classic mixer: find teammates who match each square, get their initials, race for Bingo!",
        size=16,
        color=C["muted"],
    )
    add_round_rect(s, Inches(0.55), Inches(2.0), Inches(5.9), Inches(4.6), C["white"], C["line"])
    add_textbox(s, Inches(0.8), Inches(2.2), Inches(5.4), Inches(0.4), "How to play", size=18, bold=True, color=C["teal"])
    add_numbered(
        s,
        Inches(0.8),
        Inches(2.75),
        Inches(5.4),
        Inches(3.6),
        [
            "Print or share a Bingo card (next slide).",
            "People walk around (or use chat / breakouts) and find someone who fits each square.",
            "That person writes their initials — one person can only fill ONE square per card.",
            "First to complete a line (row / column / diagonal) shouts BINGO!",
            "Optional: full card = Blackout Bingo for a bigger prize.",
        ],
        size=14,
        spacing=8,
    )
    add_round_rect(s, Inches(6.8), Inches(2.0), Inches(5.9), Inches(4.6), C["sand"])
    add_textbox(s, Inches(7.05), Inches(2.2), Inches(5.4), Inches(0.4), "Facilitator notes", size=18, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.05),
        Inches(2.75),
        Inches(5.4),
        Inches(3.6),
        [
            "Prep: 1 card per person + pens (or shared Miro/Excel).",
            "Teams of any size; best with 8+ people.",
            "Verify winners by calling out a few initials.",
            "Remote: paste card in chat; mark names in a shared sheet.",
            "Prize idea: choose next month's game theme.",
        ],
        size=14,
        spacing=10,
    )
    footer(s, 4)

    # 5. Bingo card
    s = new_slide(prs, blank)
    section_label(s, "Game 01  ·  Card")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.65),
        Inches(12),
        Inches(0.5),
        "Human Bingo — sample card",
        size=28,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    bingo_cells = [
        "Has a pet",
        "Speaks 2+\nlanguages",
        "Loves spicy\nfood",
        "Morning\nperson",
        "Has visited\n3+ countries",
        "Plays an\ninstrument",
        "Left-handed",
        "Coffee\naddict",
        "Can cook a\nsignature dish",
        "Born same\nmonth as you",
        "Funny\nnickname",
        "Works from\nanother city",
        "FREE\nSPACE",
        "Binge-watched\nlast week",
        "Prefers tea\nover coffee",
        "Has a named\nplant",
        "Knows a\nmagic trick",
        "Loves\nkaraoke",
        "Can whistle\nloudly",
        "Night owl",
        "Owns 50+\nbooks",
        "Has a twin\nor sibling",
        "Can solve a\nRubik's cube",
        "Has broken\na bone",
        "Ran / walked\n5K recently",
    ]
    cell_w, cell_h = Inches(2.2), Inches(1.05)
    start_x, start_y = Inches(1.15), Inches(1.35)
    for i, text in enumerate(bingo_cells):
        r, c = divmod(i, 5)
        x = start_x + c * cell_w
        y = start_y + r * cell_h
        is_free = text.startswith("FREE")
        fill = C["coral"] if is_free else C["white"]
        add_rect(s, x, y, cell_w, cell_h, fill, C["teal"] if not is_free else C["coral"])
        add_textbox(
            s,
            x + Inches(0.08),
            y + Inches(0.2),
            cell_w - Inches(0.16),
            cell_h - Inches(0.25),
            text,
            size=12,
            bold=is_free,
            color=C["white"] if is_free else C["ink"],
            align=PP_ALIGN.CENTER,
        )
    footer(s, 5)

    # 6. Two Truths
    s = new_slide(prs, blank)
    section_label(s, "Game 02  ·  15–25 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Two Truths and a Lie", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.4),
        "Each person shares three statements. The group votes on which one is the lie.",
        size=16,
        color=C["muted"],
    )
    add_round_rect(s, Inches(0.55), Inches(1.9), Inches(6.0), Inches(4.7), C["white"], C["line"])
    add_textbox(s, Inches(0.8), Inches(2.1), Inches(5.5), Inches(0.4), "How to play", size=18, bold=True, color=C["teal"])
    add_numbered(
        s,
        Inches(0.8),
        Inches(2.65),
        Inches(5.5),
        Inches(3.7),
        [
            "Everyone secretly prepares 2 true facts + 1 lie about themselves.",
            "One person reads all three (no hints in tone).",
            "Group discusses briefly, then votes (raise hands / Zoom poll).",
            "Reveal the lie. Correct guessers get 1 point; speaker gets 1 if most are fooled.",
            "Rotate until time is up or everyone has had a turn.",
        ],
        size=14,
        spacing=10,
    )
    add_round_rect(s, Inches(6.85), Inches(1.9), Inches(5.9), Inches(4.7), C["sand"])
    add_textbox(s, Inches(7.1), Inches(2.1), Inches(5.4), Inches(0.4), "Tips & variations", size=18, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.1),
        Inches(2.65),
        Inches(5.4),
        Inches(3.7),
        [
            "Keep statements work-safe and fun.",
            "Variation: Two Lies and a Truth for harder mode.",
            "Variation: team vs team — score collectively.",
            "Remote: use anonymous chat votes or Mentimeter.",
            "Icebreaker prompt list on next slide for shy groups.",
        ],
        size=14,
        spacing=12,
    )
    footer(s, 6)

    # 7. Two Truths examples
    s = new_slide(prs, blank)
    section_label(s, "Game 02  ·  Prompts")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.7),
        Inches(12),
        Inches(0.55),
        "Example statements (for inspiration)",
        size=28,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.35),
        Inches(12),
        Inches(0.4),
        "Facilitator can model first. Use your own truths — these are samples only.",
        size=14,
        color=C["muted"],
    )
    examples = [
        ("A", "I once met a celebrity in an elevator.", "I can name every capital in Europe.", "I have skydived."),
        ("B", "I speak three languages.", "I hate chocolate.", "I lived on a boat for a month."),
        ("C", "I have never broken a bone.", "I was on TV as a kid.", "I collect vintage stamps."),
    ]
    for i, (label, t1, t2, t3) in enumerate(examples):
        y = Inches(1.9 + i * 1.6)
        add_round_rect(s, Inches(0.55), y, Inches(12.2), Inches(1.4), C["white"], C["line"])
        add_textbox(s, Inches(0.8), y + Inches(0.25), Inches(1), Inches(0.4), f"Set {label}", size=16, bold=True, color=C["coral"])
        add_textbox(s, Inches(2.0), y + Inches(0.25), Inches(10.4), Inches(1.0), f"1) {t1}\n2) {t2}\n3) {t3}", size=15, color=C["ink"])
    footer(s, 7)

    # 8. Trivia intro
    s = new_slide(prs, blank)
    section_label(s, "Game 03  ·  20–30 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Team Trivia Sprint", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.4),
        "Small teams race through themed rounds. Fast answers, loud cheers, official answer key included.",
        size=16,
        color=C["muted"],
    )
    add_round_rect(s, Inches(0.55), Inches(1.9), Inches(6.0), Inches(4.7), C["white"], C["line"])
    add_textbox(s, Inches(0.8), Inches(2.1), Inches(5.5), Inches(0.4), "How to play", size=18, bold=True, color=C["teal"])
    add_numbered(
        s,
        Inches(0.8),
        Inches(2.65),
        Inches(5.5),
        Inches(3.7),
        [
            "Split into teams of 3–5. Pick fun team names.",
            "Host reads a question; teams write answers on paper / chat.",
            "30–45 seconds per question. No phones!",
            "Reveal answer, award 1 point (or 2 for hard Qs).",
            "3 rounds x 5 questions. Highest score wins.",
        ],
        size=14,
        spacing=12,
    )
    add_round_rect(s, Inches(6.85), Inches(1.9), Inches(5.9), Inches(4.7), C["sand"])
    add_textbox(s, Inches(7.1), Inches(2.1), Inches(5.4), Inches(0.4), "Suggested rounds", size=18, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.1),
        Inches(2.65),
        Inches(5.4),
        Inches(3.7),
        [
            "Round 1: General Knowledge (easy warm-up)",
            "Round 2: Movies & Music",
            "Round 3: Science & Weird Facts",
            "Bonus: Lightning round — 10 yes/no in 60 sec",
            "Optional company round: swap in your org trivia",
        ],
        size=14,
        spacing=12,
    )
    footer(s, 8)

    # 9. Trivia questions
    s = new_slide(prs, blank)
    section_label(s, "Game 03  ·  Questions")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.65),
        Inches(12),
        Inches(0.45),
        "Trivia questions — hide answers until the end!",
        size=26,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    rounds = [
        (
            "Round 1 · General",
            [
                "What is the largest planet in our solar system?",
                "How many continents are there on Earth?",
                "What language has the most native speakers?",
                "Which element has the chemical symbol Au?",
                "What year did the first iPhone launch?",
            ],
        ),
        (
            "Round 2 · Movies & Music",
            [
                "Who directed the movie Inception?",
                "Which band sang Bohemian Rhapsody?",
                "What is the name of the coffee shop in Friends?",
                "Which Disney film features a talking snowman named Olaf?",
                "What instrument has 88 keys?",
            ],
        ),
        (
            "Round 3 · Science & Weird",
            [
                "How many bones are in an adult human body?",
                "What is the fastest land animal?",
                "True or false: bananas are berries (botanically).",
                "What gas do plants absorb that humans exhale?",
                "About how long does sunlight take to reach Earth?",
            ],
        ),
    ]
    for i, (title, qs) in enumerate(rounds):
        x = Inches(0.45 + i * 4.25)
        add_round_rect(s, x, Inches(1.25), Inches(4.05), Inches(5.5), C["white"], C["line"])
        add_textbox(s, x + Inches(0.2), Inches(1.4), Inches(3.65), Inches(0.45), title, size=15, bold=True, color=C["coral"])
        box = s.shapes.add_textbox(x + Inches(0.2), Inches(2.0), Inches(3.65), Inches(4.5))
        tf = box.text_frame
        tf.word_wrap = True
        for j, q in enumerate(qs):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.space_after = Pt(14)
            run = p.add_run()
            run.text = f"{j + 1}. {q}"
            set_run(run, size=12, color=C["ink"])
    footer(s, 9)

    # 10. Trivia answers
    s = new_slide(prs, blank)
    section_label(s, "Game 03  ·  Answer key  ·  Facilitator only")
    add_textbox(s, Inches(0.6), Inches(0.65), Inches(12), Inches(0.45), "Trivia solutions", size=28, bold=True, color=C["ink"], font_name="Georgia")
    answers = [
        ("Round 1", ["Jupiter", "7", "Mandarin Chinese", "Gold", "2007"]),
        ("Round 2", ["Christopher Nolan", "Queen", "Central Perk", "Frozen", "Piano"]),
        (
            "Round 3",
            [
                "206",
                "Cheetah",
                "True (bananas are berries; strawberries are not)",
                "Carbon dioxide (CO2)",
                "About 8 minutes (~8 min 20 sec)",
            ],
        ),
    ]
    for i, (title, ans) in enumerate(answers):
        x = Inches(0.45 + i * 4.25)
        add_round_rect(s, x, Inches(1.3), Inches(4.05), Inches(5.4), C["sand"])
        add_textbox(s, x + Inches(0.2), Inches(1.5), Inches(3.65), Inches(0.4), title, size=16, bold=True, color=C["teal"])
        box = s.shapes.add_textbox(x + Inches(0.2), Inches(2.1), Inches(3.65), Inches(4.3))
        tf = box.text_frame
        tf.word_wrap = True
        for j, a in enumerate(ans):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.space_after = Pt(16)
            run = p.add_run()
            run.text = f"{j + 1}. {a}"
            set_run(run, size=14, bold=True, color=C["ink"])
    footer(s, 10)

    # 11. Scavenger hunt
    s = new_slide(prs, blank)
    section_label(s, "Game 04  ·  15–25 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Photo Scavenger Hunt", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.4),
        "Teams complete creative photo / video challenges. Best for hybrid & remote alike.",
        size=16,
        color=C["muted"],
    )
    add_round_rect(s, Inches(0.55), Inches(1.85), Inches(5.9), Inches(4.75), C["white"], C["line"])
    add_textbox(s, Inches(0.8), Inches(2.05), Inches(5.4), Inches(0.4), "How to play", size=18, bold=True, color=C["teal"])
    add_numbered(
        s,
        Inches(0.8),
        Inches(2.6),
        Inches(5.4),
        Inches(3.7),
        [
            "Share the challenge list (next slide).",
            "Teams have 12–15 minutes to capture as many as possible.",
            "Submit photos to a shared album / Slack channel.",
            "Host scores: 1 pt per item, +2 for creativity, +3 for funniest.",
            "Showcase top 5 photos on screen and crown winners.",
        ],
        size=14,
        spacing=10,
    )
    add_round_rect(s, Inches(6.8), Inches(1.85), Inches(5.9), Inches(4.75), C["sand"])
    add_textbox(s, Inches(7.05), Inches(2.05), Inches(5.4), Inches(0.4), "Scoring rubric", size=18, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.05),
        Inches(2.6),
        Inches(5.4),
        Inches(3.7),
        [
            "Completed item: 1 point",
            "Whole team in frame: +1",
            "Most creative interpretation: +2",
            "Crowd-favorite laugh: +3",
            "Penalty: none — encourage wild ideas!",
        ],
        size=14,
        spacing=12,
    )
    footer(s, 11)

    # 12. Scavenger list
    s = new_slide(prs, blank)
    section_label(s, "Game 04  ·  Challenge list")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.7),
        Inches(12),
        Inches(0.5),
        "Photo challenges (pick 8–10)",
        size=28,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    challenges = [
        "Something older than 20 years",
        "All teammates wearing the same color",
        "A book with a weird title",
        "A plant + a human posing like it",
        "The most chaotic desk setup",
        "A pet (or stuffed animal cameo)",
        "Recreate a famous movie poster",
        "An optical illusion with household items",
        "A homemade trophy for your team",
        "The silliest synchronized pose",
        "A window view from two different homes",
        "Spell your team name with objects",
    ]
    for i, text in enumerate(challenges):
        col, row = i % 3, i // 3
        x = Inches(0.5 + col * 4.2)
        y = Inches(1.4 + row * 1.35)
        add_round_rect(s, x, y, Inches(4.0), Inches(1.15), C["white"], C["line"])
        add_textbox(s, x + Inches(0.25), y + Inches(0.35), Inches(3.5), Inches(0.5), f"{i + 1}.  {text}", size=13, color=C["ink"])
    footer(s, 12)

    # 13. Emoji Charades
    s = new_slide(prs, blank)
    section_label(s, "Game 05  ·  15–20 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Emoji Charades", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.4),
        "Act out or emoji-spell a prompt. Teammates guess the movie, song, or workplace phrase.",
        size=16,
        color=C["muted"],
    )
    add_round_rect(s, Inches(0.55), Inches(1.85), Inches(6.0), Inches(4.75), C["white"], C["line"])
    add_textbox(s, Inches(0.8), Inches(2.05), Inches(5.5), Inches(0.4), "How to play", size=18, bold=True, color=C["teal"])
    add_numbered(
        s,
        Inches(0.8),
        Inches(2.6),
        Inches(5.5),
        Inches(3.7),
        [
            "Split into 2 teams. One actor per turn.",
            "Actor draws a prompt (no speaking the title).",
            "Mode A: silent acting (classic charades).",
            "Mode B: only emoji in chat / whiteboard.",
            "60–90 seconds to guess. Correct = 1 point.",
        ],
        size=14,
        spacing=12,
    )
    add_round_rect(s, Inches(6.85), Inches(1.85), Inches(5.9), Inches(4.75), C["sand"])
    add_textbox(s, Inches(7.1), Inches(2.05), Inches(5.4), Inches(0.4), "Prompt bank (title = answer)", size=16, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.1),
        Inches(2.6),
        Inches(5.4),
        Inches(3.7),
        [
            "The Lion King",
            "Titanic",
            "Finding Nemo",
            "Spider-Man",
            "Happy Birthday",
            "Standup that could have been an email",
            "Mute yourself on Zoom",
            "Coffee break",
            "Ship it Friday",
            "Out of office",
        ],
        size=14,
        spacing=6,
    )
    footer(s, 13)

    # 14. Minute challenges
    s = new_slide(prs, blank)
    section_label(s, "Game 06  ·  15–20 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Minute Challenges", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.4),
        "One-minute mini-games. Rotate stations or play one-at-a-time with a shared timer.",
        size=16,
        color=C["muted"],
    )
    challenges6 = [
        ("Stack Attack", "Stack as many sticky notes / cups as you can in 60 seconds without falling."),
        ("Emoji Story", "Write a 1-sentence story using only emoji. Team decodes it."),
        ("Paper Plane", "Build a plane and throw for distance (hallway / Zoom measure)."),
        ("Name Chain", "In a circle, recite everyone's name + a fruit in order — restart on fail."),
        ("Desk Stretch", "Follow host's 60-sec stretch; vote funniest form."),
        ("Blind Sketch", "Draw a simple object while eyes closed; team guesses."),
    ]
    for i, (t, d) in enumerate(challenges6):
        col, row = i % 3, i // 3
        x = Inches(0.5 + col * 4.2)
        y = Inches(1.9 + row * 2.4)
        add_round_rect(s, x, y, Inches(4.0), Inches(2.15), C["white"], C["line"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.25), Inches(3.6), Inches(0.4), t, size=16, bold=True, color=C["coral"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.8), Inches(3.6), Inches(1.1), d, size=13, color=C["ink"])
    footer(s, 14)

    # 15. Would You Rather
    s = new_slide(prs, blank)
    section_label(s, "Game 07  ·  10–15 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Would You Rather", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.35),
        "Two options. Pick a side. Defend your choice in 20 seconds. Great cooler / closer.",
        size=16,
        color=C["muted"],
    )
    wyr = [
        "Always be 10 minutes early  OR  always be exactly on time?",
        "Work from a beach for a month  OR  a cabin in the mountains?",
        "Only communicate via memes  OR  only via voice notes?",
        "Have unlimited coffee  OR  unlimited snacks at work?",
        "Fight one horse-sized duck  OR  100 duck-sized horses?",
        "Always know when someone is lying  OR  always get away with lying?",
        "Meet your future self for 10 min  OR  your past self for 10 min?",
        "Be the funniest person in the room  OR  the smartest?",
    ]
    for i, q in enumerate(wyr):
        col, row = i % 2, i // 2
        x = Inches(0.5 + col * 6.35)
        y = Inches(1.85 + row * 1.2)
        add_round_rect(s, x, y, Inches(6.15), Inches(1.05), C["white"], C["line"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.3), Inches(5.75), Inches(0.55), f"{i + 1}.  {q}", size=13, color=C["ink"])
    footer(s, 15)

    # 16. Office Olympics
    s = new_slide(prs, blank)
    section_label(s, "Game 08  ·  20–30 min")
    add_textbox(s, Inches(0.6), Inches(0.7), Inches(12), Inches(0.55), "Office Olympics", size=36, bold=True, color=C["ink"], font_name="Georgia")
    add_textbox(
        s,
        Inches(0.6),
        Inches(1.3),
        Inches(12),
        Inches(0.35),
        "A relay of silly desk sports. Best in-person; remote teams can adapt with household items.",
        size=16,
        color=C["muted"],
    )
    olympics = [
        ("Cotton Ball Race", "Move cotton balls with a spoon from bowl A to B. First team done wins."),
        ("Sticky Note Ski", "Slide on sticky-note skis (smooth floor only — safety first)."),
        ("Rubber Band Archery", "Shoot rubber bands at a paper target. Most hits wins."),
        ("Chair Spin Memory", "Spin 3x in a chair, then recall a 5-item list the host showed."),
        ("Paperclip Chain", "Link the longest paperclip chain in 2 minutes."),
        ("Silent Opera", "Sing a work slogan as opera — judges score drama."),
    ]
    for i, (t, d) in enumerate(olympics):
        col, row = i % 3, i // 3
        x = Inches(0.5 + col * 4.2)
        y = Inches(1.85 + row * 2.45)
        fill = C["sand"] if row == 0 else C["white"]
        add_round_rect(s, x, y, Inches(4.0), Inches(2.2), fill, C["line"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.25), Inches(3.6), Inches(0.45), t, size=15, bold=True, color=C["teal"])
        add_textbox(s, x + Inches(0.2), y + Inches(0.85), Inches(3.6), Inches(1.15), d, size=13, color=C["ink"])
    footer(s, 16)

    # 17. Sample plan
    s = new_slide(prs, blank)
    section_label(s, "Ready-made agenda")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.7),
        Inches(12),
        Inches(0.5),
        "Sample 60-minute event",
        size=32,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    plan = [
        ("0–5 min", "Welcome + team energy check", "Host intro, kindness rules, prize tease"),
        ("5–20 min", "Human Bingo", "Icebreaker — get people talking"),
        ("20–40 min", "Team Trivia Sprint", "Main event — 3 rounds + answers"),
        ("40–55 min", "Emoji Charades or Photo Hunt", "High-energy closer"),
        ("55–60 min", "Awards + shout-outs", "Winners, MVP teammate, next month tease"),
    ]
    for i, (when, what, why) in enumerate(plan):
        y = Inches(1.4 + i * 1.0)
        add_round_rect(s, Inches(0.55), y, Inches(12.2), Inches(0.85), C["white"], C["line"])
        add_rect(s, Inches(0.55), y, Inches(0.15), Inches(0.85), C["coral"] if i % 2 == 0 else C["teal"])
        add_textbox(s, Inches(0.95), y + Inches(0.22), Inches(2.2), Inches(0.45), when, size=14, bold=True, color=C["teal"])
        add_textbox(s, Inches(3.3), y + Inches(0.22), Inches(4.5), Inches(0.45), what, size=15, bold=True, color=C["ink"])
        add_textbox(s, Inches(8.0), y + Inches(0.22), Inches(4.4), Inches(0.45), why, size=13, color=C["muted"])
    footer(s, 17)

    # 18. Solutions sheet
    s = new_slide(prs, blank)
    section_label(s, "Facilitator solutions sheet")
    add_textbox(
        s,
        Inches(0.6),
        Inches(0.65),
        Inches(12),
        Inches(0.45),
        "Answer keys & judging notes",
        size=28,
        bold=True,
        color=C["ink"],
        font_name="Georgia",
    )
    add_round_rect(s, Inches(0.5), Inches(1.25), Inches(6.0), Inches(5.5), C["white"], C["line"])
    add_textbox(s, Inches(0.75), Inches(1.45), Inches(5.5), Inches(0.4), "Trivia (quick ref)", size=16, bold=True, color=C["coral"])
    add_bullets(
        s,
        Inches(0.75),
        Inches(2.0),
        Inches(5.5),
        Inches(2.0),
        [
            "R1: Jupiter · 7 · Mandarin · Gold · 2007",
            "R2: Nolan · Queen · Central Perk · Frozen · Piano",
            "R3: 206 · Cheetah · True · CO2 · ~8 minutes",
        ],
        size=13,
        spacing=8,
    )
    add_textbox(s, Inches(0.75), Inches(3.5), Inches(5.5), Inches(0.4), "Two Truths sample sets", size=16, bold=True, color=C["coral"])
    add_bullets(
        s,
        Inches(0.75),
        Inches(4.05),
        Inches(5.5),
        Inches(2.4),
        [
            "These are EXAMPLES only — use your own truths!",
            "Decide which statement is the lie before you model.",
            "Never force personal reveals; allow a pass.",
        ],
        size=13,
        spacing=8,
    )
    add_round_rect(s, Inches(6.8), Inches(1.25), Inches(5.95), Inches(5.5), C["sand"])
    add_textbox(s, Inches(7.05), Inches(1.45), Inches(5.5), Inches(0.4), "Charades prompts = answers", size=16, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.05),
        Inches(2.0),
        Inches(5.5),
        Inches(2.2),
        [
            "Movie/song prompts: the title is the answer.",
            "Workplace phrases: exact phrase wins; close counts.",
        ],
        size=13,
        spacing=8,
    )
    add_textbox(s, Inches(7.05), Inches(3.5), Inches(5.5), Inches(0.4), "Bingo verification", size=16, bold=True, color=C["teal"])
    add_bullets(
        s,
        Inches(7.05),
        Inches(4.05),
        Inches(5.5),
        Inches(2.2),
        [
            "Ask winner to read 3 squares + who signed.",
            "Spot-check with those people.",
            "Ties: first Blackout, or shared prize.",
        ],
        size=13,
        spacing=8,
    )
    footer(s, 18)

    # 19. Closing
    s = new_slide(prs, blank, C["teal_dark"])
    add_rect(s, 0, Inches(0.12), Inches(13.333), Inches(7.38), C["teal_dark"])
    add_textbox(
        s,
        Inches(0.8),
        Inches(2.2),
        Inches(11.5),
        Inches(1.0),
        "Have fun out there.",
        size=40,
        bold=True,
        color=C["white"],
        font_name="Georgia",
    )
    add_textbox(
        s,
        Inches(0.8),
        Inches(3.3),
        Inches(11),
        Inches(1.0),
        "Swap in company-specific trivia, print Bingo cards, and keep prizes\nlight. The goal is connection — not perfection.",
        size=18,
        color="D0E8E9",
    )
    add_textbox(
        s,
        Inches(0.8),
        Inches(5.2),
        Inches(11),
        Inches(0.5),
        "Reuse this deck every month — rotate games, refresh prompts.",
        size=15,
        color="A8D5D7",
    )

    out_repo = "/agent/repos/course/team-events/Monthly_Team_Event_Games_Playbook.pptx"
    out_art = "/opt/cursor/artifacts/Monthly_Team_Event_Games_Playbook.pptx"
    prs.save(out_repo)
    prs.save(out_art)
    print(f"Saved: {out_repo}")
    print(f"Artifacts: {out_art}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
