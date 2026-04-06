from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

W, H = 1920, 1080
ROOT = Path(__file__).resolve().parent
SLIDES = ROOT / 'boardroom_slides'
LOGOS = ROOT / 'logos'
SLIDES.mkdir(parents=True, exist_ok=True)

FONT_REG = '/System/Library/Fonts/Supplemental/Avenir Next.ttc'
FONT_COND = '/System/Library/Fonts/Supplemental/Avenir Next Condensed.ttc'

BG = '#f6f8fb'
BG2 = '#f6f8fb'
PANEL = '#ffffff'
PANEL2 = '#ffffff'
CREAM = '#12253d'
TEXT = '#39506b'
MUTED = '#60758d'
MUTED2 = '#8b9bb0'
ACCENT = '#0b6bcb'
ACCENT2 = '#12806a'
GOLD = '#b7791f'
ORANGE = '#d97706'
RED = '#c2410c'
LINE = '#d8e1ec'


def font(path, size):
    return ImageFont.truetype(path, size)

F_KICK = font(FONT_REG, 20)
F_TITLE = font(FONT_COND, 72)
F_SUB = font(FONT_REG, 24)
F_H2 = font(FONT_COND, 40)
F_H3 = font(FONT_COND, 28)
F_BODY = font(FONT_REG, 22)
F_BODY_SM = font(FONT_REG, 18)
F_LABEL = font(FONT_REG, 18)
F_SMALL = font(FONT_REG, 16)
F_CHIP = font(FONT_REG, 19)
F_SUB_LG = font(FONT_REG, 28)
F_LABEL_LG = font(FONT_REG, 22)
F_BODY_MID = font(FONT_REG, 20)
F_CHIP_SM = font(FONT_REG, 17)


def hexrgba(h, a=255):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4)) + (a,)


def bg():
    im = Image.new('RGBA', (W, H), BG)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, W, 12), fill=hexrgba(ACCENT))
    d.rectangle((W - 380, 0, W, H), fill=hexrgba('#eef3f8'))
    return im


def shadow_card(base, xy, radius=24, fill=PANEL, outline=LINE, blur=8, offset=(0, 6), shadow=(29, 52, 84, 18)):
    ov = Image.new('RGBA', base.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    x1, y1, x2, y2 = xy
    od.rounded_rectangle((x1 + offset[0], y1 + offset[1], x2 + offset[0], y2 + offset[1]), radius=radius, fill=shadow)
    d = ImageDraw.Draw(base)
    base.alpha_composite(ov)
    d.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=2)


def wrap(draw, text, fnt, width):
    words = text.split()
    lines = []
    cur = ''
    for w in words:
        test = w if not cur else cur + ' ' + w
        if draw.textlength(test, font=fnt) <= width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return '\n'.join(lines)


def add_header(draw, title, subtitle, x=120, y=94, width=1180, sub_font=F_SUB):
    title_wrapped = wrap(draw, title, F_TITLE, width)
    draw.multiline_text((x, y), title_wrapped, font=F_TITLE, fill=CREAM, spacing=6)
    tb = draw.multiline_textbbox((x, y), title_wrapped, font=F_TITLE, spacing=6)
    suby = tb[3] + 36
    subtitle_wrapped = wrap(draw, subtitle, sub_font, width)
    draw.multiline_text((x, suby), subtitle_wrapped, font=sub_font, fill=MUTED, spacing=10)
    sb = draw.multiline_textbbox((x, suby), subtitle_wrapped, font=sub_font, spacing=10)
    return sb[3]


def logo_chip(base, draw, label, logo_name, x, y, w=None, h=58, fill='#ffffff', font_obj=F_CHIP, logo_box=None):
    logo = Image.open(LOGOS / logo_name).convert('RGBA')
    max_logo = logo_box or (h - 18)
    logo.thumbnail((max_logo, max_logo))
    tw = draw.textlength(label, font=font_obj)
    if w is None:
        w = int(44 + logo.width + tw)
    shadow_card(base, (x, y, x + w, y + h), radius=20, fill=fill, outline=LINE, blur=8, offset=(0, 4), shadow=(29, 52, 84, 16))
    base.alpha_composite(logo, (x + 16, y + (h - logo.height) // 2))
    text_y = y + (h - (font_obj.size if hasattr(font_obj, "size") else 20)) // 2 - 1
    draw.text((x + 28 + logo.width, text_y), label, font=font_obj, fill=CREAM)


def text_chip(base, draw, label, x, y, w=None, h=58, fill='#ffffff', outline=LINE, font_obj=F_CHIP, align='left'):
    tw = draw.textlength(label, font=font_obj)
    if w is None:
        w = int(tw + 44)
    shadow_card(base, (x, y, x + w, y + h), radius=20, fill=fill, outline=outline, blur=8, offset=(0, 4), shadow=(29, 52, 84, 16))
    text_y = y + (h - (font_obj.size if hasattr(font_obj, "size") else 20)) // 2 - 1
    if align == 'center':
        text_x = x + (w - tw) / 2
    else:
        text_x = x + 22
    draw.text((text_x, text_y), label, font=font_obj, fill=CREAM)


def bullet_list(draw, items, x, y, width, bullet=GOLD, font_obj=F_BODY, gap=16):
    yy = y
    for item in items:
        draw.ellipse((x, yy + 8, x + 8, yy + 16), fill=hexrgba(bullet))
        wrapped = wrap(draw, item, font_obj, width - 32)
        draw.multiline_text((x + 24, yy), wrapped, font=font_obj, fill=CREAM, spacing=6)
        bb = draw.multiline_textbbox((x + 24, yy), wrapped, font=font_obj, spacing=6)
        yy = bb[3] + gap
    return yy


def tile(base, draw, x, y, w, h, title, body, logo_name=None, accent=ACCENT, title_font=F_H3, body_font=F_BODY_SM, logo_box=40):
    shadow_card(base, (x, y, x + w, y + h), radius=28, fill=PANEL)
    if logo_name:
        logo = Image.open(LOGOS / logo_name).convert('RGBA')
        logo.thumbnail((logo_box, logo_box))
        base.alpha_composite(logo, (x + 24, y + 24))
        tx = x + 80
    else:
        tx = x + 24
    draw.text((tx, y + 22), title, font=title_font, fill=CREAM)
    draw.rounded_rectangle((tx, y + 72, tx + 110, y + 78), radius=3, fill=hexrgba(accent))
    draw.multiline_text((x + 24, y + 96), wrap(draw, body, body_font, w - 48), font=body_font, fill=TEXT, spacing=8)


def harvey_screenshot_panel(base, draw, x, y, w, h):
    shadow_card(base, (x, y, x + w, y + h), radius=22, fill='#f3f5f8', outline=LINE, blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
    logo = Image.open(LOGOS / 'harvey.png').convert('RGBA')
    logo.thumbnail((32, 32))
    base.alpha_composite(logo, (x + 26, y + 20))
    draw.text((x + 68, y + 22), 'Harvey', font=F_LABEL, fill=MUTED)
    # prompt bar
    draw.rounded_rectangle((x + 18, y + 70, x + w - 16, y + 158), radius=18, fill='#ffffff', outline=LINE)
    prompt = 'Can you take this PDF and remove the signature pages to create a signature packet?'
    draw.multiline_text((x + 54, y + 86), wrap(draw, prompt, F_BODY_SM, w - 108), font=F_BODY_SM, fill='#1f2937', spacing=6)
    draw.ellipse((x + 32, y + 91, x + 44, y + 103), fill=hexrgba(MUTED2))
    # toolbar / status
    draw.text((x + 32, y + 176), 'Copy      Save prompt      Edit query', font=F_SMALL, fill=MUTED)
    draw.text((x + 32, y + 224), 'Finished in 3 steps', font=F_BODY_SM, fill='#1f2937')
    draw.text((x + 230, y + 224), '›', font=F_BODY_SM, fill=MUTED)
    # response body
    body = (
        'I appreciate the request, but I\'m not able to create, edit, or modify PDF files or any other file types. '
        'I can only read and analyze the content of uploaded documents, not produce new versions of them.'
    )
    draw.multiline_text((x + 32, y + 288), wrap(draw, body, F_BODY_SM, w - 64), font=F_BODY_SM, fill='#334155', spacing=10)


def slide_1():
    im = bg(); d = ImageDraw.Draw(im)
    draw_kick = 'KIRKLAND AI × EMMANEIGH'
    d.rounded_rectangle((120, 70, 440, 116), radius=18, fill='#eef3f8', outline=LINE, width=2)
    d.text((146, 85), draw_kick, font=F_KICK, fill=MUTED)
    d.multiline_text((120, 160), 'The model can reason.\nEmmaNeigh makes it execute.', font=F_TITLE, fill=CREAM, spacing=6)
    d.multiline_text((120, 448), wrap(d, 'A model-agnostic last-mile adoption layer: it gets firm documents into AI-usable form and executes work across the legal stack.', F_SUB_LG, 820), font=F_SUB_LG, fill=MUTED, spacing=10)
    shadow_card(im, (1035, 118, 1785, 842), radius=28, fill=PANEL2)
    d.text((1090, 164), 'Execution layer', font=F_H2, fill=CREAM)
    d.multiline_text((1090, 234), wrap(d, 'EmmaNeigh sits between the model and the operating stack. It turns intent into executable commands and routes the work across the tools lawyers already use.', F_BODY, 620), font=F_BODY, fill=TEXT, spacing=8)
    chip_gap = 16
    harvey_w = 178
    firm_w = 236
    open_w = 156
    logo_chip(im, d, 'Harvey', 'harvey.png', 1090, 366, w=harvey_w, h=52, font_obj=F_CHIP_SM)
    text_chip(im, d, 'Firm-approved model', 1090 + harvey_w + chip_gap, 366, w=firm_w, h=52, font_obj=F_CHIP_SM)
    text_chip(im, d, 'Open model', 1090 + harvey_w + chip_gap + firm_w + chip_gap, 366, w=open_w, h=52, font_obj=F_CHIP_SM, align='center')
    bottom_chip_gap = 20
    future_w = 276
    managed_w = 292
    text_chip(im, d, 'Future internal model', 1090, 428, w=future_w, h=52, font_obj=F_CHIP_SM)
    text_chip(im, d, 'Managed remote model', 1090 + future_w + bottom_chip_gap, 428, w=managed_w, h=52, font_obj=F_CHIP_SM)
    shadow_card(im, (1090, 534, 1710, 616), radius=22, fill='#edf5ff', outline='#90bff0', blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
    d.text((1124, 560), 'EmmaNeigh execution fabric', font=F_H3, fill=CREAM)
    xx = 1120
    yy = 670
    tools = [('Outlook', 'outlook.png', 190), ('Word', 'word.png', 150), ('Adobe', 'adobe.png', 175), ('DocuSign', 'docusign.png', 200), ('Litera', 'litera_favicon.png', 150), ('iManage', 'imanage.png', 180)]
    for label, logo, ww in tools[:3]:
        logo_chip(im, d, label, logo, xx, yy, w=ww, h=46)
        xx += ww + 18
    xx = 1120
    yy = 730
    for label, logo, ww in tools[3:]:
        logo_chip(im, d, label, logo, xx, yy, w=ww, h=46)
        xx += ww + 18
    return im


def slide_2():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'The last mile is adoption across the desktop stack.', 'Cloud AI can reason about prompts. It still cannot reliably access local files, convert firm documents into AI-usable inputs, or execute across desktop applications and matter systems.', 120, 86, 1260)
    d.text((120, bottom + 42), 'Representative matter stack', font=F_LABEL_LG, fill=MUTED)
    tiles = [
        ('Outlook', 'Search matter traffic, save attachments, confirm who sent what', 'outlook.png', ACCENT),
        ('Litera', 'Run redlines and produce comparison output', 'litera_favicon.png', ORANGE),
        ('iManage', 'Browse, save, version, and organize matter files', 'imanage.png', ACCENT2),
        ('Adobe', 'Clean PDFs, remove pages, combine sets, and prep packets', 'adobe.png', RED),
        ('DocuSign', 'Circulate signature pages and rebuild executed versions', 'docusign.png', '#8e7dff'),
        ('Word', 'Prepare clean documents and convert deliverables', 'word.png', '#7aa5ff'),
    ]
    x0, y0 = 120, bottom + 82
    col_w, row_h, gap_x, gap_y = 270, 178, 18, 18
    for i, (t, b, logo, accent) in enumerate(tiles):
        r, c = divmod(i, 3)
        tile(im, d, x0 + c * (col_w + gap_x), y0 + r * (row_h + gap_y), col_w, row_h, t, b, logo, accent, body_font=F_BODY_MID)
    harvey_x = 1060
    harvey_w = 700
    harvey_screenshot_panel(im, d, harvey_x, bottom + 68, harvey_w, 520)
    d.rounded_rectangle((harvey_x, bottom + 612, harvey_x + harvey_w, bottom + 722), radius=18, fill='#fff7e8', outline=hexrgba('#edd6ab'), width=2)
    d.multiline_text((harvey_x + 30, bottom + 638), wrap(d, 'Key point: the model is not refusing because it does not understand the task. It is refusing because it does not have the operating-system and application-level execution surface.', F_BODY_SM, harvey_w - 60), font=F_BODY_SM, fill=hexrgba(GOLD), spacing=6)
    return im


def slide_3():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'EmmaNeigh closes the adoption gap between reasoning and execution.', 'First it gets matter content into a form the model can interact with. Then it routes the work across connected software and files.', 120, 86, 1240)
    shadow_card(im, (120, bottom + 70, 1800, 942), radius=26, fill=PANEL2)
    d.text((170, bottom + 106), '1. Reasoning layer', font=F_H2, fill=CREAM)
    d.multiline_text((170, bottom + 166), wrap(d, 'Harvey or another model interprets the request, identifies the workflow, and fills in the parameters.', F_BODY, 420), font=F_BODY, fill=TEXT, spacing=8)
    logo_chip(im, d, 'Harvey', 'harvey.png', 170, bottom + 274, w=170, h=50)
    text_chip(im, d, 'Other model', 170, bottom + 334, w=210, h=50)
    d.text((760, bottom + 120), '2. EmmaNeigh execution layer', font=F_H2, fill=CREAM)
    shadow_card(im, (760, bottom + 194, 1740, bottom + 354), radius=22, fill='#edf5ff', outline='#90bff0', blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
    d.multiline_text((800, bottom + 222), wrap(d, 'Turns local documents, PDFs, metadata, and natural-language intent into machine-readable commands and deterministic workflows.', F_H3, 860), font=F_H3, fill='#163962', spacing=4)
    d.text((760, bottom + 378), '3. Operating systems and applications', font=F_H2, fill=CREAM)
    tools = [('Local files', 'adobe.png', 180), ('Outlook', 'outlook.png', 180), ('Word', 'word.png', 150), ('Adobe', 'adobe.png', 165), ('Litera', 'litera_favicon.png', 150), ('DocuSign', 'docusign.png', 190), ('iManage', 'imanage.png', 220)]
    xx, yy = 760, bottom + 446
    for label, logo, ww in tools:
        logo_chip(im, d, label, logo, xx, yy, w=ww, h=58)
        xx += ww + 14
        if xx > 1630:
            xx = 760
            yy += 76
    bullet_list(d, [
        'Last-mile adoption requires two things: getting matter content into an AI-usable format and executing the resulting workflow across real systems.',
        'Execution itself does not require AI. AI is the interface layer; EmmaNeigh handles translation and system-level execution.'
    ], 170, bottom + 418, 540, bullet=GOLD, font_obj=F_BODY_SM, gap=12)
    return im


def slide_4():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'What EmmaNeigh already operationalizes', 'These are last-mile adoption workflows: getting matter content into usable form and carrying the work through the systems where it actually happens.', 120, 86, 1180)
    cards = [
        ('Checklist updates', 'Upload a checklist, scan matter activity, and update comments and status.', 'outlook.png', ACCENT),
        ('Punchlists', 'Turn a working checklist into a cleaner matter-management punchlist.', 'outlook.png', '#7aa5ff'),
        ('Redlines', 'Run Litera comparisons and generate comparison outputs.', 'litera_favicon.png', ORANGE),
        ('PDF workflows', 'Remove signature pages, combine sets, and prep execution-ready PDF materials.', 'adobe.png', RED),
        ('Executed versions', 'Match signed pages back into agreements and rebuild executed sets.', 'docusign.png', '#8e7dff'),
        ('Document movement', 'Move, convert, save, and organize deliverables across files and connected systems.', 'imanage.png', ACCENT2),
    ]
    x0, y0 = 120, bottom + 70
    col_w, row_h, gap_x, gap_y = 530, 190, 24, 24
    for i, (title, body, logo, accent) in enumerate(cards):
        r, c = divmod(i, 2)
        logo_box = 52 if title == 'Document movement' else 40
        tile(im, d, x0 + c * (col_w + gap_x), y0 + r * (row_h + gap_y), col_w, row_h, title, body, logo, accent, logo_box=logo_box)
    shadow_card(im, (1248, y0, 1800, y0 + 610), radius=30, fill=PANEL2)
    d.text((1290, y0 + 42), 'Operating principle', font=F_H2, fill=CREAM)
    bullet_list(d, [
        'These workflows can be invoked manually or via AI routing.',
        'The same workflow engine can work beneath Harvey today and another reasoning layer later.'
    ], 1290, y0 + 152, 450, bullet=ACCENT, font_obj=F_H3, gap=30)
    return im


def slide_5():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'Why firms care', 'This is exactly where adoption stalls today: matter administration and paralegal work that still has to be carried through local systems.', 120, 86, 1200, sub_font=F_SUB_LG)
    boxes = [
        ('Matter administration', 'Checklists, packet assembly, version handling, distribution tracking, and status chasing are necessary but operational.', GOLD),
        ('Paralegal workflows', 'Much of the work is rules-based, repetitive, and document-centric. It does not require high-end legal reasoning to execute correctly.', ACCENT),
        ('Partner write-offs', 'When lawyers do these tasks anyway, the time is hard to bill, easy to discount, and distracting from higher-value work.', ORANGE),
    ]
    x = 120
    for title, body, accent in boxes:
        shadow_card(im, (x, bottom + 80, x + 530, bottom + 470), radius=30, fill=PANEL)
        d.text((x + 34, bottom + 116), title, font=F_H2, fill=CREAM)
        d.rounded_rectangle((x + 34, bottom + 192, x + 164, bottom + 200), radius=4, fill=hexrgba(accent))
        d.multiline_text((x + 34, bottom + 226), wrap(d, body, F_BODY_MID, 460), font=F_BODY_MID, fill=TEXT, spacing=10)
        x += 570
    shadow_card(im, (120, bottom + 530, 1800, 950), radius=30, fill=PANEL2)
    statement = 'EmmaNeigh does not replace legal judgment. It compresses the operational execution time around documents, signatures, redlines, checklists, PDFs, and matter administration so lawyers spend less time on work clients resist paying for.'
    d.multiline_text((160, bottom + 570), wrap(d, statement, F_H3, 1500), font=F_H3, fill=CREAM, spacing=10)
    return im


def slide_6():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'Why this is interesting for Kirkland AI', 'EmmaNeigh is not another model discussion. It is the last-mile adoption layer that can sit beneath the firm’s preferred reasoning system.', 120, 86, 1360)
    shadow_card(im, (120, bottom + 70, 940, 900), radius=32, fill=PANEL2)
    d.text((160, bottom + 120), 'Kirkland AI implication', font=F_H2, fill=CREAM)
    d.multiline_text((160, bottom + 185), wrap(d, 'If Kirkland wants to push Harvey or another reasoning system beyond drafting and analysis, EmmaNeigh is the layer that can make those systems usable across the existing desktop and DMS environment.', F_BODY, 700), font=F_BODY, fill=TEXT, spacing=8)
    d.text((160, bottom + 360), 'Illustrative stack', font=F_LABEL, fill=MUTED)
    for i, (label, logo, ww) in enumerate([('Harvey', 'harvey.png', 170), ('Outlook', 'outlook.png', 180), ('Adobe', 'adobe.png', 165), ('Litera', 'litera_favicon.png', 150), ('DocuSign', 'docusign.png', 195), ('iManage', 'imanage.png', 220)]):
        row = 0 if i < 3 else 1
        x = 160 + (i % 3) * 230
        y = bottom + 410 + row * 82
        logo_chip(im, d, label, logo, x, y, w=ww, h=58)
    shadow_card(im, (1000, bottom + 70, 1800, 900), radius=32, fill=PANEL)
    d.text((1040, bottom + 120), 'What Kirkland gets if this works', font=F_H2, fill=CREAM)
    bullet_list(d, [
        'A model-agnostic control plane that can work with Harvey today and a different reasoning system tomorrow.',
        'A path from AI analysis into concrete execution across the existing desktop and DMS stack.',
        'Operational leverage without requiring lawyers to manually bridge the gap between cloud AI and local systems.',
        'No rip-and-replace: the product sits on top of the tools the firm already uses.'
    ], 1040, bottom + 220, 680, bullet=GOLD, font_obj=F_BODY, gap=18)
    return im


def slide_7():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'Two inputs would materially improve the product at Kirkland.', 'These are not cosmetic asks. They are the two leverage points that would make last-mile adoption materially more seamless.', 120, 86, 1280)
    shadow_card(im, (120, bottom + 80, 910, 850), radius=32, fill=PANEL)
    d.text((160, bottom + 130), '1. Harvey bearer / authentication token', font=F_H2, fill=CREAM)
    d.multiline_text((160, bottom + 220), wrap(d, 'This would allow EmmaNeigh to plug into a reasoning layer with stronger legal performance than a free foundation model while preserving the model-agnostic execution architecture.', F_BODY, 690), font=F_BODY, fill=TEXT, spacing=8)
    shadow_card(im, (160, bottom + 360, 340, bottom + 416), radius=20, fill='#171315', outline=LINE, blur=8, offset=(0, 4), shadow=(29, 52, 84, 16))
    harvey = Image.open(LOGOS / 'harvey.png').convert('RGBA')
    harvey.thumbnail((28, 28))
    im.alpha_composite(harvey, (180, bottom + 374))
    d.text((216, bottom + 374), 'Harvey', font=F_CHIP, fill='#ffffff')
    shadow_card(im, (160, bottom + 445, 860, bottom + 610), radius=20, fill='#edf5ff', outline='#90bff0', blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
    d.text((190, bottom + 462), 'Why it matters', font=F_LABEL, fill=hexrgba(ACCENT))
    d.multiline_text((190, bottom + 490), wrap(d, 'Better planning, better task routing, and better explanation quality without changing the execution engine underneath.', F_BODY, 620), font=F_BODY, fill=CREAM, spacing=8)
    shadow_card(im, (1010, bottom + 80, 1800, 850), radius=32, fill=PANEL)
    d.text((1050, bottom + 130), '2. Expanded iManage API surface', font=F_H2, fill=CREAM)
    d.multiline_text((1050, bottom + 220), wrap(d, 'Right now the accessible interface is constrained. That limits browse, version-aware workflows, and redlining across versions. A richer surface would let the product operate far more seamlessly inside the document system.', F_BODY, 690), font=F_BODY, fill=TEXT, spacing=8)
    logo_chip(im, d, 'iManage', 'imanage.png', 1050, bottom + 360, w=220, h=56)
    shadow_card(im, (1050, bottom + 445, 1750, bottom + 610), radius=20, fill='#edf5ff', outline='#90bff0', blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
    d.text((1080, bottom + 462), 'Why it matters', font=F_LABEL, fill=hexrgba(ACCENT))
    d.multiline_text((1080, bottom + 490), wrap(d, 'The product gets closer to full browse, save, version, and compare workflows instead of stopping at partial desktop integration.', F_BODY, 620), font=F_BODY, fill=CREAM, spacing=8)
    d.multiline_text((120, 972), wrap(d, 'These are the two inputs that matter most.', F_BODY_SM, 1600), font=F_BODY_SM, fill=hexrgba(GOLD), spacing=6)
    return im


def slide_8():
    im = bg(); d = ImageDraw.Draw(im)
    bottom = add_header(d, 'Moving toward a pilot', 'The strongest test is a focused pilot against a handful of real transaction workflows with the right system access and the right file/format surface.', 120, 86, 1420)
    shadow_card(im, (120, bottom + 80, 1060, 900), radius=32, fill=PANEL2)
    d.text((160, bottom + 130), 'Suggested pilot structure', font=F_H2, fill=CREAM)
    steps = [
        ('Step 1', 'Connect the reasoning layer and target integrations', ACCENT),
        ('Step 2', 'Validate three workflows end-to-end on sample matters', GOLD),
        ('Step 3', 'Run a small user group, measure cycle-time reduction, and identify which permissions unlock the next tier.', ORANGE),
    ]
    yy = bottom + 224
    for i, (wk, desc, color) in enumerate(steps):
        d.ellipse((180, yy + 4, 204, yy + 28), fill=hexrgba(color))
        if i < len(steps) - 1:
            d.line((192, yy + 30, 192, yy + 96), fill=hexrgba(color), width=5)
        d.text((240, yy - 4), wk, font=F_H3, fill=CREAM)
        d.multiline_text((240, yy + 42), wrap(d, desc, F_BODY_SM, 700), font=F_BODY_SM, fill=TEXT, spacing=8)
        yy += 128
    d.text((1160, bottom + 86), 'What success would look like', font=F_H2, fill=CREAM)
    outcomes = [
        ('Faster checklist maintenance', 'Comments and status updated from actual matter activity'),
        ('Less manual packet work', 'Signature and executed-version assembly compressed materially'),
        ('Lower write-off pressure', 'Less lawyer time spent on operational execution'),
        ('Clear integration roadmap', 'Know exactly which permissions unlock the next tier of value'),
    ]
    yy = bottom + 144
    for head, sub in outcomes:
        shadow_card(im, (1160, yy, 1740, yy + 124), radius=20, fill='#f7fbff', outline='#cfe0f4', blur=8, offset=(0, 4), shadow=(29, 52, 84, 14))
        d.text((1190, yy + 18), head, font=F_H3, fill=CREAM)
        d.multiline_text((1190, yy + 58), wrap(d, sub, F_BODY_SM, 500), font=F_BODY_SM, fill=TEXT, spacing=6)
        yy += 144
    return im


slides = [
    ('01_cover.png', slide_1),
    ('02_problem.png', slide_2),
    ('03_execution_layer.png', slide_3),
    ('04_workflows.png', slide_4),
    ('05_value.png', slide_5),
    ('06_kirkland.png', slide_6),
    ('07_asks.png', slide_7),
    ('08_pilot.png', slide_8),
]

for name, fn in slides:
    img = fn()
    img.save(SLIDES / name)
    print(SLIDES / name)
