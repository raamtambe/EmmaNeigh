from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math
import textwrap

W, H = 1920, 1080
ROOT = Path(__file__).resolve().parent
SLIDES = ROOT / 'slides'
LOGOS = ROOT / 'logos'
SLIDES.mkdir(parents=True, exist_ok=True)

FONT_REG = '/System/Library/Fonts/Supplemental/Avenir Next.ttc'
FONT_COND = '/System/Library/Fonts/Supplemental/Avenir Next Condensed.ttc'
FONT_ALT = '/System/Library/Fonts/HelveticaNeue.ttc'

BG = '#07111f'
BG2 = '#0b1730'
CARD = '#0d1b2c'
CARD2 = '#111f34'
CREAM = '#f6f1e8'
MUTED = '#a8b5c8'
MUTED2 = '#7d8ea7'
BLUE = '#38bdf8'
CYAN = '#67e8f9'
TEAL = '#2dd4bf'
PURPLE = '#8b5cf6'
ORANGE = '#fb923c'
GOLD = '#fbbf24'
RED = '#fb7185'
GREEN = '#79e0a6'
LINE = '#24364d'


def font(path, size):
    return ImageFont.truetype(path, size)

F_TITLE = font(FONT_COND, 96)
F_H1 = font(FONT_COND, 76)
F_H2 = font(FONT_COND, 50)
F_H3 = font(FONT_COND, 34)
F_SUB = font(FONT_REG, 28)
F_BODY = font(FONT_REG, 24)
F_BODY_SM = font(FONT_REG, 20)
F_LABEL = font(FONT_REG, 22)
F_SMALL = font(FONT_REG, 18)
F_CHIP = font(FONT_REG, 20)
F_METRIC = font(FONT_COND, 58)


def rr(draw, xy, r, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def shadow(base, xy, r, fill, shadow_color=(0, 0, 0, 120), blur=18, offset=(0, 10), outline=None, width=1):
    overlay = Image.new('RGBA', base.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    x1, y1, x2, y2 = xy
    od.rounded_rectangle((x1 + offset[0], y1 + offset[1], x2 + offset[0], y2 + offset[1]), radius=r, fill=shadow_color)
    overlay = overlay.filter(ImageFilter.GaussianBlur(blur))
    base.alpha_composite(overlay)
    d = ImageDraw.Draw(base)
    d.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def gradient_bg(top=BG, bottom=BG2):
    im = Image.new('RGBA', (W, H), top)
    px = im.load()
    import colorsys
    def hexrgb(h):
        h=h.lstrip('#')
        return tuple(int(h[i:i+2],16) for i in (0,2,4))
    t = hexrgb(top)
    b = hexrgb(bottom)
    for y in range(H):
        a = y / (H - 1)
        r = int(t[0]*(1-a) + b[0]*a)
        g = int(t[1]*(1-a) + b[1]*a)
        bb = int(t[2]*(1-a) + b[2]*a)
        for x in range(W):
            px[x,y] = (r,g,bb,255)
    # soft radial glows
    for cx, cy, color, radius, alpha in [
        (W*0.78, H*0.18, PURPLE, 560, 150),
        (W*0.16, H*0.82, TEAL, 620, 90),
        (W*0.88, H*0.85, BLUE, 540, 70),
    ]:
        glow = Image.new('RGBA', (W, H), (0,0,0,0))
        gd = ImageDraw.Draw(glow)
        rgb = tuple(int(color[i:i+2], 16) for i in (1,3,5))
        gd.ellipse((cx-radius, cy-radius, cx+radius, cy+radius), fill=rgb + (alpha,))
        glow = glow.filter(ImageFilter.GaussianBlur(140))
        im.alpha_composite(glow)
    return im


def add_title(draw, title, subtitle=None, x=100, y=90, width=1500):
    title_text = wrap(title, F_H1, width, draw)
    draw.multiline_text((x, y), title_text, font=F_H1, fill=CREAM, spacing=6)
    bbox = draw.multiline_textbbox((x, y), title_text, font=F_H1, spacing=6)
    if subtitle:
        subtitle_text = wrap(subtitle, F_SUB, width, draw)
        draw.multiline_text((x, bbox[3] + 28), subtitle_text, font=F_SUB, fill=MUTED, spacing=10)
        sub_bbox = draw.multiline_textbbox((x, bbox[3] + 28), subtitle_text, font=F_SUB, spacing=10)
        return sub_bbox[3]
    return bbox[3]


def wrap(text, fnt, width, draw):
    words = text.split()
    lines=[]
    cur=''
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


def bullet_list(draw, items, x, y, width, bullet_color=CYAN, fnt=F_BODY, gap=26, bullet_size=10):
    yy = y
    for item in items:
        draw.ellipse((x, yy+12, x+bullet_size, yy+12+bullet_size), fill=hex2rgba(bullet_color))
        txt = wrap(item, fnt, width-40, draw)
        draw.multiline_text((x+28, yy), txt, font=fnt, fill=CREAM, spacing=8)
        bbox = draw.multiline_textbbox((x+28, yy), txt, font=fnt, spacing=8)
        yy = bbox[3] + gap
    return yy


def hex2rgba(hexstr, alpha=255):
    h = hexstr.lstrip('#')
    return tuple(int(h[i:i+2],16) for i in (0,2,4)) + (alpha,)


def chip(draw, text, x, y, pad_x=18, pad_y=10, fill='#111f34', outline='#2a4261', text_fill=CREAM, icon=None, icon_size=28):
    tw = draw.textlength(text, font=F_CHIP)
    w = int(tw + pad_x*2 + (icon_size + 12 if icon else 0))
    h = 48
    rr(draw, (x,y,x+w,y+h), 24, fill=fill, outline=outline, width=2)
    tx = x + pad_x
    if icon is not None:
        icon.thumbnail((icon_size, icon_size))
    return w, h


def paste_logo(base, name, x, y, max_wh=(120,60), mode='contain'):
    p = LOGOS / name
    im = Image.open(p).convert('RGBA')
    if mode == 'contain':
        im.thumbnail(max_wh)
    else:
        im = im.resize(max_wh)
    base.alpha_composite(im, (x, y))
    return im.size


def pill_with_logo(base, draw, label, logo_name, x, y, w=None, h=60, fill='#0e1f32', outline='#24364d'):
    logo = Image.open(LOGOS / logo_name).convert('RGBA')
    logo.thumbnail((h-16, h-16))
    text_w = draw.textlength(label, font=F_CHIP)
    if w is None:
        w = int(36 + logo.width + 14 + text_w + 24)
    shadow(base, (x, y, x+w, y+h), 28, fill=fill, shadow_color=(0,0,0,90), blur=12, offset=(0,8), outline=outline, width=2)
    base.alpha_composite(logo, (x+18, y + (h-logo.height)//2))
    draw.text((x+18+logo.width+14, y+15), label, font=F_CHIP, fill=CREAM)
    return w


def small_card(base, draw, title, subtitle, x, y, w, h, logo_name=None):
    shadow(base, (x,y,x+w,y+h), 30, fill='#0c1a2a', shadow_color=(0,0,0,110), blur=20, offset=(0,12), outline=LINE, width=2)
    if logo_name:
        paste_logo(base, logo_name, x+22, y+20, (56,56))
    draw.text((x+92 if logo_name else x+24, y+20), title, font=F_H3, fill=CREAM)
    draw.multiline_text((x+24, y+86), wrap(subtitle, F_BODY_SM, w-48, draw), font=F_BODY_SM, fill=MUTED, spacing=8)


def window_card(base, draw, title, subtitle, x, y, w, h, logo_name=None, accent=BLUE):
    shadow(base, (x,y,x+w,y+h), 28, fill='#091423', shadow_color=(0,0,0,120), blur=24, offset=(0,14), outline='#29415f', width=2)
    rr(draw, (x+1,y+1,x+w-1,y+46), 28, fill='#101b2b')
    # top dots
    for i,c in enumerate([(255,96,92),(255,189,68),(0,202,78)]):
        draw.ellipse((x+18+i*18, y+16, x+28+i*18, y+26), fill=c)
    if logo_name:
        lg = Image.open(LOGOS/logo_name).convert('RGBA')
        lg.thumbnail((30,30))
        base.alpha_composite(lg,(x+60,y+9))
        tx=x+98
    else:
        tx=x+18
    draw.text((tx,y+11), title, font=F_LABEL, fill=CREAM)
    rr(draw, (x+22, y+68, x+w-22, y+110), 16, fill='#0f233d')
    draw.text((x+40,y+77), subtitle, font=F_BODY_SM, fill=hex2rgba(accent))
    # content lines
    yy = y+140
    for k in range(4):
        ww = w-60 if k!=3 else int((w-60)*0.62)
        rr(draw, (x+30, yy, x+30+ww, yy+18), 9, fill='#15283f')
        yy += 36


def quote_panel(base, draw, x, y, w, h):
    shadow(base, (x,y,x+w,y+h), 32, fill='#0b1322', shadow_color=(0,0,0,125), blur=24, offset=(0,18), outline='#304863', width=2)
    draw.text((x+36,y+28), 'Representative cloud-AI constraint', font=F_SMALL, fill=hex2rgba(GOLD))
    pill_with_logo(base, draw, 'Harvey / other reasoning model', 'harvey.png', x+36, y+64, h=54, fill='#111013', outline='#2c2a30')
    q = '“I can analyze this agreement, but I cannot create a sig packet, run Litera, or save a version into iManage from your desktop.”'
    draw.multiline_text((x+36,y+152), wrap(q, F_H3, w-72, draw), font=F_H3, fill=CREAM, spacing=10)
    draw.multiline_text((x+36,y+h-120), wrap('The reasoning layer is in the cloud. The work still lives across local files, desktop applications, and firm systems.', F_BODY_SM, w-72, draw), font=F_BODY_SM, fill=MUTED, spacing=8)


def slide_cover():
    im = gradient_bg()
    d = ImageDraw.Draw(im)
    # hero glow line
    rr(d, (100,120,360,160), 20, fill='#0f2240', outline='#244264', width=2)
    d.text((124,129), 'KIRKLAND AI × EMMANEIGH', font=F_LABEL, fill=MUTED)
    title = 'The model can reason.\nEmmaNeigh makes\nit execute.'
    d.multiline_text((100,220), title, font=F_TITLE, fill=CREAM, spacing=0)
    sub = 'A model-agnostic execution layer for matter administration, paralegal workflows, and document operations across the legal stack.'
    d.multiline_text((100,540), wrap(sub, F_SUB, 880, d), font=F_SUB, fill=MUTED, spacing=10)
    shadow(im, (100,690,760,820), 28, fill='#111f34', shadow_color=(0,0,0,115), blur=24, offset=(0,14), outline='#3b5e87', width=2)
    d.text((136,730), 'Execution is the moat.', font=F_H2, fill=hex2rgba(GOLD))
    d.multiline_text((136,786), wrap('Models will change. The workflow fabric compounds.', F_BODY, 560, d), font=F_BODY, fill=CREAM, spacing=6)
    # diagram right
    shadow(im, (1130,180,1770,820), 38, fill='#091321', shadow_color=(0,0,0,120), blur=30, offset=(0,18), outline='#24354a', width=2)
    d.text((1180,230), 'Execution fabric', font=F_H2, fill=CREAM)
    d.multiline_text((1180,305), wrap('AI interprets intent. EmmaNeigh routes and executes the work across the actual operating stack.', F_BODY, 520, d), font=F_BODY, fill=MUTED, spacing=8)
    # top model chips
    x=1180; y=390
    for label in ['Harvey','Claude','GPT','Groq / Qwen','Local model']:
        tw = d.textlength(label, font=F_CHIP)
        w = int(tw+42)
        rr(d,(x,y,x+w,y+46),22,fill='#101b2b',outline='#2a4261',width=2)
        d.text((x+20,y+11),label,font=F_CHIP,fill=CREAM)
        x += w + 12
        if x > 1660:
            x=1180; y+=58
    shadow(im, (1180,520,1720,610), 30, fill='#112740', shadow_color=(77,182,255,35), blur=16, offset=(0,0), outline='#59c1ff', width=2)
    d.text((1208,544), 'EmmaNeigh execution layer', font=F_H3, fill=CREAM)
    # bottom logos
    coords = [(1180,680,'outlook.png','Outlook'),(1385,680,'adobe.png','Adobe'),(1590,680,'docusign.png','DocuSign'),(1180,780,'imanage.png','iManage'),(1460,780,'litera_favicon.png','Litera')]
    for xx,yy,logo,label in coords:
        w = 220 if label=='iManage' else None
        pill_with_logo(im,d,label,logo,xx,yy,w=w,h=68,fill='#0e1d30',outline='#233852')
    d.text((100,950), 'Confidential discussion draft · April 2026', font=F_SMALL, fill=MUTED2)
    return im


def slide_problem():
    im = gradient_bg('#06101d','#0b1830')
    d = ImageDraw.Draw(im)
    top = add_title(d, 'The bottleneck is not thinking. It is execution across the desktop stack.', 'Matter administration still requires people to open files, switch systems, and complete deterministic steps one application at a time.', x=90, y=70, width=1350)
    # left side attorney desktop
    content_y = top + 70
    d.text((110,content_y), 'A single matter can require live interaction across all of this:', font=F_BODY, fill=MUTED)
    window_card(im,d,'Outlook','Find the latest draft, save attachments, confirm who sent what',110,content_y+50,360,250,'outlook.png',accent=BLUE)
    window_card(im,d,'Adobe','Split, combine, clean, and package PDFs for execution and closing',500,content_y+50,360,250,'adobe.png',accent=RED)
    window_card(im,d,'iManage','Save down, version up, browse versions, organize matter files',110,content_y+330,360,250,'imanage.png',accent=CYAN)
    window_card(im,d,'Litera + DocuSign','Run redlines, create sig packets, assemble executed pages',500,content_y+330,360,250,'docusign.png',accent=ORANGE)
    # divide line
    for yy in range(content_y+10,900,28):
        d.rounded_rectangle((920,yy,930,yy+14), radius=6, fill='#284364')
    d.text((878,915), 'SYSTEM BOUNDARY', font=F_SMALL, fill=MUTED2)
    # right quote panel
    quote_panel(im,d,980,content_y+50,840,470)
    rr(d, (90, 902, 1810, 986), 22, fill='#0b1524', outline='#2a405c', width=2)
    d.multiline_text((118,928), wrap('Result: attorneys and paralegals still do high-volume operational work manually, and partners often write the time off anyway.', F_BODY_SM, 1600, d), font=F_BODY_SM, fill=hex2rgba(GOLD), spacing=6)
    return im


def slide_layer():
    im = gradient_bg('#07101b','#08182b')
    d = ImageDraw.Draw(im)
    top = add_title(d, 'EmmaNeigh unifies the AI layer and the systems-integration layer.', 'Think of it as a USB-C hub for legal work: one execution layer that plugs models into Outlook, Adobe, Litera, iManage, DocuSign, files, and checklists.', x=90, y=80, width=1350)
    # top row models
    d.text((110, top + 45), '1. Model-agnostic reasoning layer', font=F_LABEL, fill=hex2rgba(CYAN))
    mx=110; my=top + 85
    model_labels=['Harvey','Claude','GPT','Groq / Qwen','Local model']
    arrow_centers = []
    for label in model_labels:
        tw=d.textlength(label,font=F_CHIP)
        w=int(tw+52)
        rr(d,(mx,my,mx+w,my+54),27,fill='#111b2d',outline='#29415f',width=2)
        d.text((mx+24,my+14),label,font=F_CHIP,fill=CREAM)
        arrow_centers.append(mx + (w // 2))
        mx += w + 18
    # arrows down
    arrow_top = my + 54
    for x in arrow_centers:
        d.line((x,arrow_top,x,arrow_top+40), fill=hex2rgba(CYAN), width=4)
        d.polygon((x-8,arrow_top+40,x+8,arrow_top+40,x,arrow_top+56), fill=hex2rgba(CYAN))
    # execution bar
    bar_y = arrow_top + 74
    shadow(im,(110,bar_y,1810,bar_y+130),36,fill='#10253f',shadow_color=(77,182,255,45),blur=24,offset=(0,10),outline='#5ec8ff',width=3)
    d.text((150,bar_y+37),'2. EmmaNeigh execution layer',font=F_H2,fill=CREAM)
    d.multiline_text((1030,bar_y+30),wrap('Translates natural-language intent into machine-readable commands and then executes them deterministically.', F_BODY_SM, 700, d),font=F_BODY_SM,fill=MUTED,spacing=6)
    # bottom systems
    d.text((110, bar_y+180), '3. Actual operating systems and applications', font=F_LABEL, fill=hex2rgba(ORANGE))
    y = bar_y+220
    rr(d,(110,y,280,y+74),28,fill='#0d1b2c',outline='#304661',width=2)
    d.text((140,y+24),'Local files',font=F_CHIP,fill=CREAM)
    pills=[('Outlook','outlook.png'),('Adobe','adobe.png'),('Litera','litera_favicon.png'),('DocuSign','docusign.png'),('iManage','imanage.png')]
    x_positions=[310,530,740,960,1190]
    widths=[190,170,170,205,260]
    for (label,logo),xx,ww in zip(pills,x_positions,widths):
        pill_with_logo(im,d,label,logo,xx,y,w=ww,h=74,fill='#0d1b2c',outline='#304661')
    # callouts bottom
    items=[
        'AI is used for translation and routing, not for performing the underlying deterministic work.',
        'No browser-click agent is required for core execution. The value comes from software integrations, file operations, and operating-system access.',
        'Because the layer is model agnostic, the reasoning system can change without rebuilding the workflow engine.'
    ]
    bullet_list(d, items, 110, y+82, 1540, bullet_color=GOLD, fnt=F_BODY_SM, gap=10, bullet_size=8)
    return im


def slide_workflows():
    im = gradient_bg('#07111f','#0d1830')
    d = ImageDraw.Draw(im)
    top = add_title(d, 'What EmmaNeigh already executes', 'The product is already organized around recurring legal operations, not generic chat.', x=90, y=80, width=1200)
    cards=[
        ('Email & attachments', 'Search folders, determine whether a draft was received or sent, save attachments, and prepare follow-ups.', 'outlook.png'),
        ('Checklist updates', 'Upload a checklist, scan Outlook activity, and update comments/status based on actual matter traffic.', 'outlook.png'),
        ('Punchlists', 'Turn a working checklist into a cleaner punchlist format for transaction management and follow-up.', 'imanage.png'),
        ('Redlines', 'Run Litera comparisons and output full-document or targeted comparison sets.', 'litera_favicon.png'),
        ('PDF workflows', 'Split, combine, clean, and prep signature and closing PDFs.', 'adobe.png'),
        ('Executed versions', 'Match signed pages back into the correct agreements and rebuild executed sets at scale.', 'docusign.png'),
        ('Document management', 'Browse, save, organize, and version documents where the environment allows it.', 'imanage.png'),
        ('File operations', 'Convert files, save outputs, move deliverables, and handle repetitive desktop document actions.', 'adobe.png'),
    ]
    x0,y0 = 90,top+60
    w,h = 410,180
    gapx,gapy = 36,34
    for idx,(title,sub,logo) in enumerate(cards):
        row, col = divmod(idx,4)
        x = x0 + col*(w+gapx)
        y = y0 + row*(h+gapy)
        small_card(im,d,title,sub,x,y,w,h,logo)
    rr(d,(90,890,1830,990),26,fill='#0c1828',outline='#273d58',width=2)
    d.multiline_text((122,914),wrap('Core point: the execution layer is useful even without AI. AI simply makes it accessible through natural-language prompts instead of menus and macros.', F_BODY, 1580, d), font=F_BODY, fill=CREAM, spacing=6)
    return im


def slide_value():
    im = gradient_bg('#08111b','#0c172a')
    d=ImageDraw.Draw(im)
    top = add_title(d, 'Why firms care: this is matter administration and paralegal work that leaks into lawyer time.', None, x=90, y=80, width=1600)
    cols=[
        ('Matter administration','Checklists, packet assembly, version handling, distribution tracking, and status chasing are necessary but operational.', GOLD),
        ('Paralegal workflows','Much of the work is rules-based, repetitive, and document-centric. It does not require high-end legal reasoning to execute correctly.', CYAN),
        ('Partner write-offs','When lawyers do these tasks anyway, the time is hard to bill, easy to discount, and distracting from higher-value work.', ORANGE),
    ]
    x=90
    for title,sub,color in cols:
        shadow(im,(x,top+50,x+540,top+490),34,fill='#0d1a2c',shadow_color=(0,0,0,110),blur=20,offset=(0,14),outline='#273d58',width=2)
        d.text((x+36,top+98), title, font=F_H2, fill=CREAM)
        d.rectangle((x+36,top+188,x+180,top+196), fill=hex2rgba(color))
        d.multiline_text((x+36,top+236), wrap(sub, F_BODY, 450, d), font=F_BODY, fill=MUTED, spacing=10)
        x += 590
    rr(d,(90,top+540,1830,1010),34,fill='#0e2137',outline='#2b4668',width=2)
    d.text((124,top+596), 'The pitch to the firm is simple:', font=F_LABEL, fill=hex2rgba(CYAN))
    d.multiline_text((124,top+638), wrap('EmmaNeigh does not replace legal judgment. It compresses the operational execution time around documents, signatures, redlines, checklists, and matter administration so lawyers spend less time on work clients resist paying for.', F_BODY, 1560, d), font=F_BODY, fill=CREAM, spacing=8)
    return im


def slide_kirkland():
    im = gradient_bg('#07101d','#0b1730')
    d=ImageDraw.Draw(im)
    top = add_title(d, 'Why this is interesting for Kirkland AI', 'EmmaNeigh is not another standalone model. It is the execution surface that can sit beneath the reasoning layer Kirkland chooses to use.', x=90, y=80, width=1450)
    # left diagram
    left_y = top + 50
    shadow(im,(90,left_y,1040,860),36,fill='#0c1728',shadow_color=(0,0,0,115),blur=24,offset=(0,14),outline='#273d58',width=2)
    d.text((130,left_y+40),'Reasoning system',font=F_LABEL,fill=hex2rgba(CYAN))
    pill_with_logo(im,d,'Harvey','harvey.png',130,left_y+80,w=260,h=74,fill='#111013',outline='#2d2a30')
    for i,label in enumerate(['Other frontier model','Open model','Future internal model']):
        rr(d,(420,left_y+85+i*86,730,left_y+142+i*86),28,fill='#121d2f',outline='#324964',width=2)
        d.text((446,left_y+101+i*86),label,font=F_CHIP,fill=CREAM)
    for yy in [left_y+270,left_y+345,left_y+420]:
        d.line((310,left_y+170,310,yy),fill=hex2rgba(CYAN),width=4)
        d.line((310,yy,565,yy),fill=hex2rgba(CYAN),width=4)
    shadow(im,(130,left_y+310,980,left_y+420),30,fill='#112740',shadow_color=(77,182,255,45),blur=22,offset=(0,8),outline='#5ec8ff',width=2)
    d.text((160,left_y+342),'EmmaNeigh execution layer',font=F_H3,fill=CREAM)
    d.text((160,left_y+475),'Firm operating stack',font=F_LABEL,fill=hex2rgba(ORANGE))
    xx=130
    for label,logo,w in [('Outlook','outlook.png',190),('Adobe','adobe.png',170),('Litera','litera_favicon.png',170),('iManage','imanage.png',240)]:
        pill_with_logo(im,d,label,logo,xx,left_y+510,w=w,h=70,fill='#0e1d30',outline='#304661')
        xx += w + 20
    # right bullets
    shadow(im,(1120,left_y,1830,860),36,fill='#0c1728',shadow_color=(0,0,0,115),blur=24,offset=(0,14),outline='#273d58',width=2)
    d.multiline_text((1160,left_y+50),wrap('What Kirkland gets if this works', F_H2, 560, d), font=F_H2, fill=CREAM, spacing=0)
    items=[
        'A model-agnostic control plane that can work with Harvey today and a different reasoning system tomorrow.',
        'A way to extend the AI discussion from drafting and analysis into concrete execution across the existing desktop and DMS stack.',
        'A path to operational leverage without requiring lawyers to manually bridge the gap between cloud AI and local applications.',
        'No rip-and-replace: the product is designed to sit on top of the tools the firm already uses.'
    ]
    bullet_list(d, items, 1160, left_y+170, 600, bullet_color=GOLD, fnt=F_BODY, gap=16, bullet_size=10)
    return im


def slide_asks():
    im = gradient_bg('#07111f','#0c1830')
    d=ImageDraw.Draw(im)
    top = add_title(d, 'Two asks would materially improve the product in a Kirkland environment.', 'These are not cosmetic asks. They unlock the exact execution depth that the current stack otherwise blocks.', x=90, y=80, width=1450)
    card_y = top + 50
    shadow(im,(90,card_y,900,840),38,fill='#0c1828',shadow_color=(0,0,0,120),blur=26,offset=(0,16),outline='#273d58',width=2)
    d.multiline_text((130,card_y+40),wrap('1. Harvey bearer / authentication token', F_H2, 660, d),font=F_H2,fill=CREAM,spacing=0)
    d.multiline_text((130,card_y+150), wrap('That would allow EmmaNeigh to plug into a reasoning layer with stronger legal performance than a free foundation model, while preserving the model-agnostic execution architecture.', F_BODY, 690, d), font=F_BODY, fill=MUTED, spacing=10)
    rr(d,(130,card_y+370,860,card_y+500),28,fill='#10253f',outline='#5ec8ff',width=2)
    d.text((160,card_y+405),'Why it matters',font=F_LABEL,fill=hex2rgba(CYAN))
    d.multiline_text((160,card_y+443), wrap('Better planning, better task routing, better explanation quality — without changing the execution engine underneath.', F_BODY_SM, 660, d), font=F_BODY_SM, fill=CREAM, spacing=8)
    pill_with_logo(im,d,'Harvey','harvey.png',130,card_y+300,w=240,h=64,fill='#111013',outline='#2d2a30')

    shadow(im,(1020,card_y,1830,840),38,fill='#0c1828',shadow_color=(0,0,0,120),blur=26,offset=(0,16),outline='#273d58',width=2)
    d.multiline_text((1060,card_y+40),wrap('2. Unrestricted iManage API surface', F_H2, 640, d),font=F_H2,fill=CREAM,spacing=0)
    d.multiline_text((1060,card_y+150), wrap('Right now the accessible interface is constrained. That limits browse, version-aware workflows, and redlining across versions. A richer surface would let the product operate far more seamlessly inside the document system.', F_BODY, 690, d), font=F_BODY, fill=MUTED, spacing=10)
    rr(d,(1060,card_y+370,1790,card_y+500),28,fill='#10253f',outline='#5ec8ff',width=2)
    d.text((1090,card_y+405),'Why it matters',font=F_LABEL,fill=hex2rgba(CYAN))
    d.multiline_text((1090,card_y+443), wrap('The product gets closer to full browse / save / version / compare workflows instead of stopping at partial desktop integration.', F_BODY_SM, 660, d), font=F_BODY_SM, fill=CREAM, spacing=8)
    pill_with_logo(im,d,'iManage','imanage.png',1060,card_y+300,w=280,h=64,fill='#0e1d30',outline='#304661')

    d.multiline_text((90,920),wrap('If Kirkland AI wants to evaluate whether this can become a real operating layer rather than just another demo, these are the leverage points.', F_BODY, 1700, d), font=F_BODY, fill=hex2rgba(GOLD), spacing=6)
    return im


def slide_pilot():
    im = gradient_bg('#08111f','#0c172f')
    d=ImageDraw.Draw(im)
    top = add_title(d, 'A sensible next step is a narrow pilot, not a broad rollout.', 'The product is strongest when evaluated against a handful of real transaction workflows with the right system access.', x=90, y=80, width=1450)
    # left timeline
    left_y = top + 50
    shadow(im,(90,left_y,1180,960),38,fill='#0c1828',shadow_color=(0,0,0,120),blur=26,offset=(0,16),outline='#273d58',width=2)
    d.text((130,left_y+50),'Suggested pilot structure', font=F_H2, fill=CREAM)
    steps=[
        ('Week 1', 'Connect model layer + target integrations', CYAN),
        ('Week 2', 'Validate 3 workflows end-to-end on sample matters', GOLD),
        ('Weeks 3–5', 'Run a small user group, measure cycle-time reduction, and identify the permissions that unlock the next tier of value.', ORANGE),
    ]
    yy=left_y+150
    for i,(wk,desc,color) in enumerate(steps):
        d.ellipse((150,yy-2,174,yy+22), fill=hex2rgba(color))
        if i < len(steps)-1:
            d.line((162,yy+24,162,yy+96), fill=hex2rgba(color), width=5)
        d.text((210,yy-14), wk, font=F_H3, fill=CREAM)
        d.multiline_text((210,yy+42), wrap(desc, F_BODY_SM, 790, d), font=F_BODY_SM, fill=MUTED, spacing=8)
        yy += 132
    # right outcomes
    shadow(im,(1240,left_y,1830,960),38,fill='#0c1828',shadow_color=(0,0,0,120),blur=26,offset=(0,16),outline='#273d58',width=2)
    d.multiline_text((1280,left_y+50),wrap('What success would look like', F_H2, 500, d), font=F_H2, fill=CREAM, spacing=0)
    metrics=[
        ('Faster checklist maintenance', 'Comments and status updated from actual matter activity'),
        ('Less manual packet work', 'Signature and executed-version assembly compressed materially'),
        ('Lower write-off pressure', 'Less lawyer time spent on operational execution'),
        ('Clear integration roadmap', 'Know exactly which permissions unlock the next tier of value'),
    ]
    yy=left_y+150
    for m,sub in metrics:
        rr(d,(1280,yy,1780,yy+90),22,fill='#10253f',outline='#2d4668',width=2)
        d.text((1308,yy+18),m,font=F_H3,fill=CREAM)
        d.multiline_text((1308,yy+50),wrap(sub,F_BODY_SM,430,d),font=F_BODY_SM,fill=MUTED,spacing=4)
        yy += 108
    return im


slides = [
    ('01_cover.png', slide_cover),
    ('02_problem.png', slide_problem),
    ('03_execution_layer.png', slide_layer),
    ('04_workflows.png', slide_workflows),
    ('05_value.png', slide_value),
    ('06_kirkland.png', slide_kirkland),
    ('07_asks.png', slide_asks),
    ('08_pilot.png', slide_pilot),
]

for name, fn in slides:
    img = fn()
    img.save(SLIDES / name)
    print(SLIDES / name)
