"""OBD AI Solutions: client kickoff deck (16:9), generated with python-pptx.

Uses Playfair Display + Inter (bundled in assets/fonts). Install them on the
presenting machine for the PPTX to look exactly like the PDF export.
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]

ACCENT = RGBColor(0x6A, 0x26, 0xD8)
BLUE = RGBColor(0x2B, 0x3B, 0xB8)
DEEP = RGBColor(0x12, 0x0E, 0x2E)
LILAC = RGBColor(0xC9, 0xB4, 0xFF)
INK = RGBColor(0x14, 0x14, 0x14)
TEXT = RGBColor(0x2A, 0x2A, 0x2D)
MUTED = RGBColor(0x6B, 0x68, 0x78)
LINE = RGBColor(0xE6, 0xE1, 0xF0)
BLUSH = RGBColor(0xF8, 0xF6, 0xFD)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SERIF = "Playfair Display"
SANS = "Inter"

W, H = Inches(13.333), Inches(7.5)
MX = Inches(0.8)  # side margin


def money(v, cur="PKR"):
    return f"{cur} {v:,.0f}"


# ---------------------------------------------------------------- primitives
def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s


def text(slide, x, y, w, h, runs, size=14, color=TEXT, font=SANS, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None, line_spacing=1.1):
    """runs: str | list of paragraphs; a paragraph is str or list of (text, {overrides})."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = runs if isinstance(runs, list) else [runs]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if spacing:
            p.space_after = Pt(spacing)
        segs = para if isinstance(para, list) else [(para, {})]
        for seg, o in segs:
            r = p.add_run()
            r.text = seg
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.color.rgb = o.get("color", color)
    return tb


def eyebrow(slide, x, y, label, color=ACCENT, w=Inches(8)):
    text(slide, x, y, w, Inches(0.3), label.upper(), size=10, color=color, bold=True)


def accent(slide, x, y, color=ACCENT):
    rect(slide, x, y, Inches(0.35), Pt(2.4), fill=color)
    rect(slide, x + Inches(0.35), y, Inches(0.25), Pt(2.4), fill=BLUE if color == ACCENT else color)


def logo(slide, x, y, dark=False, size=Inches(0.52)):
    mark = ROOT / "assets" / "logo" / ("obd-mark-white.png" if dark else "obd-mark.png")
    slide.shapes.add_picture(str(mark), x, y, height=size)
    wx = x + Emu(int(size * 1.03)) + Inches(0.14)
    text(slide, wx, y - Emu(int(size * .02)), Inches(2.5), Inches(0.3),
         "OBD", size=int(15 * size / Inches(0.52)), font=SERIF, bold=True, color=WHITE if dark else INK)
    text(slide, wx, y + Emu(int(size * .58)), Inches(2.5), Inches(0.2),
         "A I   S O L U T I O N S", size=7, bold=True, color=LILAC if dark else ACCENT)


def picture(slide, path, x, y, w, h):
    """Place an image filling w×h, centre-cropped."""
    from PIL import Image
    iw, ih = Image.open(path).size
    pic = slide.shapes.add_picture(str(path), x, y, w, h)
    box, img = w / h, iw / ih
    if img > box:
        cut = (1 - box / img) / 2
        pic.crop_left = pic.crop_right = cut
    else:
        cut = (1 - img / box) / 2
        pic.crop_top = pic.crop_bottom = cut
    return pic


def chrome(slide, cfg, n, total, section):
    """Standard header/footer for content slides."""
    logo(slide, MX, Inches(0.42), size=Inches(0.42))
    text(slide, W - MX - Inches(4), Inches(0.52), Inches(4), Inches(0.3), section.upper(),
         size=9, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)
    rect(slide, MX, H - Inches(0.62), W - 2 * MX, Pt(0.6), fill=LINE)
    text(slide, MX, H - Inches(0.5), Inches(6), Inches(0.3),
         f"{cfg['company']['name']}  ·  {cfg['project']['code']}", size=8.5, color=MUTED)
    text(slide, W - MX - Inches(2), H - Inches(0.5), Inches(2), Inches(0.3),
         f"{n:02d} / {total:02d}", size=8.5, color=MUTED, align=PP_ALIGN.RIGHT)


def title(slide, kicker, heading, y=Inches(1.25), italic_tail=None, width=Inches(11.5)):
    eyebrow(slide, MX, y, kicker)
    runs = [(heading, {})]
    if italic_tail:
        runs.append((italic_tail, {"italic": True, "color": ACCENT}))
    text(slide, MX, y + Inches(0.32), width, Inches(0.9), [runs], size=34, font=SERIF, color=INK)
    accent(slide, MX, y + Inches(1.18))


def card(slide, x, y, w, h, fill=WHITE, border=LINE):
    return rect(slide, x, y, w, h, fill=fill, line=border, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)


def bullets(slide, x, y, w, items, size=13, color=TEXT, mark="✓", mark_color=ACCENT, gap=Inches(0.42)):
    for i, it in enumerate(items):
        yy = y + gap * i
        text(slide, x, yy, Inches(0.3), Inches(0.3), mark, size=size, color=mark_color, bold=True)
        text(slide, x + Inches(0.32), yy, w - Inches(0.32), gap, it, size=size, color=color)


# ---------------------------------------------------------------- slides
def s_title(prs, cfg):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cl = cfg["client"]
    rect(s, 0, 0, W, H, fill=WHITE)
    band = Inches(4.7)
    rect(s, 0, 0, band, H, fill=RGBColor(0, 0, 0))
    picture(s, ROOT / "assets" / "logo" / "source-mark-circuit.jpg", 0, 0, band, Inches(5.6))
    rect(s, 0, Inches(5.6), band, H - Inches(5.6), fill=DEEP)
    text(s, Inches(0.6), Inches(5.95), band - Inches(1.2), Inches(0.6),
         [[("Build.  ", {}), ("Launch.  ", {}), ("Support.", {"color": LILAC})]], size=22, font=SERIF, color=WHITE)
    text(s, Inches(0.6), Inches(6.6), band - Inches(1.0), Inches(0.4), cfg["company"]["tagline"], size=9, color=LILAC)

    x = band + Inches(0.9)
    logo(s, x, Inches(0.7), size=Inches(0.48))
    eyebrow(s, x, Inches(1.9), f"Project kickoff  ·  {cl['business']}")
    text(s, x, Inches(2.3), Inches(7.6), Inches(2.2),
         [[(f"{cl['business']}, your website", {})], [("is almost ready to launch.", {"italic": True, "color": ACCENT})]],
         size=42, font=SERIF, color=INK, line_spacing=1.0)
    accent(s, x, Inches(4.2))
    text(s, x, Inches(4.5), Inches(7.2), Inches(0.8),
         f"All {len(cl['site_pages'])} pages are built and getting their final touches. Next, we take it live on your own domain "
         "with a professional mailbox, at no deployment cost.", size=14, color=MUTED, line_spacing=1.3)
    rect(s, x, Inches(6.05), Inches(7.2), Pt(0.6), fill=LINE)
    meta = [("Prepared for", f"{cl['business']}, {cfg['project']['city']}"), ("Presented by", cfg["company"]["founder"]),
            ("Date", cfg["project"]["issue_date"])]
    for i, (k, v) in enumerate(meta):
        xx = x + Inches(2.45) * i
        text(s, xx, Inches(6.2), Inches(2.3), Inches(0.25), k.upper(), size=8, color=MUTED, bold=True)
        text(s, xx, Inches(6.45), Inches(2.3), Inches(0.4), v, size=12, color=INK)


def s_about(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Who we are")
    title(s, "01 · OBD AI Solutions", "An engineering studio ", italic_tail="built on AI.")
    text(s, MX, Inches(2.75), Inches(6.2), Inches(1.4),
         "We design and ship AI products, MVPs and production web platforms for founders and growing "
         "businesses, working with clients in the United States, Australia and Pakistan.",
         size=15, color=MUTED, line_spacing=1.35)
    cr = cfg["credentials"]
    stats = [(cr["international_clients"], "US & Australian\nclients served"),
             (cr["local_clients"], "Businesses served\nin Pakistan"),
             ("3", "Countries\nUS · AU · PK")]
    for i, (big, lab) in enumerate(stats):
        x = MX + Inches(2.1) * i
        text(s, x, Inches(4.35), Inches(2), Inches(0.9), big, size=46, font=SERIF, color=ACCENT)
        text(s, x, Inches(5.3), Inches(2), Inches(0.7), lab, size=11, color=MUTED, line_spacing=1.2)

    services = [("AI Solutions", "Intelligent automation, LLM assistants and data workflows."),
                ("MVP Development", "Idea to investor-ready product, on foundations that scale."),
                ("Web Engineering", "Fast, secure websites with infrastructure, email and support.")]
    x0, cw = Inches(7.55), Inches(4.98)
    for i, (h, d) in enumerate(services):
        y = Inches(2.7) + Inches(1.28) * i
        card(s, x0, y, cw, Inches(1.1), fill=BLUSH if i != 2 else WHITE, border=ACCENT if i == 2 else LINE)
        text(s, x0 + Inches(0.3), y + Inches(0.17), Inches(0.6), Inches(0.4), f"0{i+1}", size=11, bold=True, color=ACCENT)
        text(s, x0 + Inches(0.8), y + Inches(0.15), cw - Inches(1), Inches(0.4), h, size=16, font=SERIF, color=INK)
        text(s, x0 + Inches(0.8), y + Inches(0.55), cw - Inches(1), Inches(0.5), d, size=11, color=MUTED)
    text(s, x0 + Inches(0.8), Inches(6.55), cw, Inches(0.3), "← Your project", size=9, bold=True, color=ACCENT)


def s_experience(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Experience")
    title(s, "02 · Track record", "The teams we've ", italic_tail="built with.")
    names = cfg["credentials"]["notable"]
    cw, gap = Inches(2.78), Inches(0.13)
    for i, nm in enumerate(names):
        x = MX + (cw + gap) * i
        card(s, x, Inches(2.85), cw, Inches(1.55), fill=BLUSH)
        parts = nm.split(" (")
        text(s, x + Inches(0.15), Inches(3.1), cw - Inches(0.3), Inches(0.75), parts[0], size=17, font=SERIF, color=INK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
        if len(parts) > 1:
            text(s, x, Inches(3.95), cw, Inches(0.3), parts[1].rstrip(")").upper(), size=9, bold=True,
                 color=ACCENT, align=PP_ALIGN.CENTER)
    text(s, MX, Inches(4.7), Inches(11.7), Inches(0.4),
         f"…plus {cfg['credentials']['local_clients']} local businesses across Pakistan.",
         size=13, color=MUTED, align=PP_ALIGN.CENTER)
    # principles strip
    rect(s, MX, Inches(5.3), W - 2 * MX, Inches(1.2), fill=DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
    pr = [("Founder-led", "Omer oversees every launch"), ("Documented", "Scope, fees & support in writing"),
          ("Transparent", "No markups on third-party costs"), ("Yours", "Everything registered to you")]
    for i, (h, d) in enumerate(pr):
        x = MX + Inches(0.4) + Inches(2.93) * i
        text(s, x, Inches(5.5), Inches(2.8), Inches(0.4), h, size=15, font=SERIF, color=WHITE)
        text(s, x, Inches(5.9), Inches(2.8), Inches(0.4), d, size=10.5, color=LILAC)


def s_team(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Your team")
    title(s, "03 · Who you'll work with", "Your ", italic_tail="project team.")
    people = [(cfg["company"]["founder"], "Project Lead", [cfg["company"]["founder_title"], cfg["company"]["founder_secondary"]],
               "Owns technical decisions, setup, configuration and launch. Your escalation point for anything critical."),
              (cfg["project"]["account_manager"], "Account Manager", ["Client relations, OBD AI Solutions"],
               "Your day-to-day contact for scheduling, documents, billing and updates.")]
    for i, (nm, role, lines, desc) in enumerate(people):
        x = MX + Inches(5.95) * i
        card(s, x, Inches(2.85), Inches(5.75), Inches(3.4), fill=WHITE, border=ACCENT if i == 0 else LINE)
        initials = "".join(w[0] for w in nm.split()[:3] if w[0].isupper())
        c = rect(s, x + Inches(0.4), Inches(3.25), Inches(1.0), Inches(1.0), fill=ACCENT if i == 0 else INK, shape=MSO_SHAPE.OVAL)
        tf = c.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = initials
        r.font.name, r.font.size, r.font.color.rgb = SERIF, Pt(18), WHITE
        eyebrow(s, x + Inches(1.65), Inches(3.3), role, w=Inches(3.8))
        text(s, x + Inches(1.65), Inches(3.58), Inches(3.9), Inches(0.5), nm, size=22, font=SERIF, color=INK)
        text(s, x + Inches(1.65), Inches(4.08), Inches(3.9), Inches(0.6), lines, size=10.5, color=MUTED, line_spacing=1.2)
        rect(s, x + Inches(0.4), Inches(4.9), Inches(4.95), Pt(0.6), fill=LINE)
        text(s, x + Inches(0.4), Inches(5.05), Inches(4.95), Inches(1), desc, size=12, color=TEXT, line_spacing=1.3)


def s_status(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cl = cfg["client"]
    chrome(s, cfg, n, t, "Your website")
    title(s, f"04 · {cl['business']} website", "Built from your content. ", italic_tail="Almost ready.")
    pages = cl["site_pages"]
    cw = (W - 2 * MX - Inches(0.15) * 3) / 4
    for i, pg in enumerate(pages + ["Almost ready"]):
        x = MX + (cw + Inches(0.15)) * (i % 4)
        y = Inches(2.75) + Inches(0.95) * (i // 4)
        last = i == len(pages)
        card(s, x, y, cw, Inches(0.8), fill=DEEP if last else (WHITE if i == 0 else BLUSH),
             border=DEEP if last else (ACCENT if i == 0 else LINE))
        text(s, x + Inches(0.25), y + Inches(0.12), cw, Inches(0.25), "STATUS" if last else f"PAGE {i+1:02d}",
             size=8, bold=True, color=LILAC if last else ACCENT)
        text(s, x + Inches(0.25), y + Inches(0.34), cw - Inches(0.4), Inches(0.4), pg, size=15, font=SERIF,
             color=WHITE if last else INK)
    y = Inches(4.85)
    text(s, MX, y, Inches(6), Inches(0.3), "RECEIVED FROM YOU, THANK YOU", size=9, bold=True, color=ACCENT)
    for i, item in enumerate(cl["received_content"]):
        text(s, MX, y + Inches(0.35) + Inches(0.36) * i, Inches(0.3), Inches(0.3), "✓", size=12, color=ACCENT, bold=True)
        text(s, MX + Inches(0.32), y + Inches(0.35) + Inches(0.36) * i, Inches(6), Inches(0.3), item, size=12.5, color=TEXT)
    card(s, MX + Inches(6.9), y, Inches(4.83), Inches(1.45), fill=BLUSH)
    text(s, MX + Inches(7.2), y + Inches(0.2), Inches(4.3), Inches(1.1),
         [[("All content is in. ", {"bold": True, "color": INK})],
          "We're adding the final touches now, and the site will be ready for your review shortly."], size=12.5, color=MUTED, line_spacing=1.3)


def s_roadmap(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Launch roadmap")
    title(s, "05 · Timeline", "A clear path ", italic_tail="to go live.")
    steps = [("Step 1", "Agreement", "Sign the Service Agreement and settle the development fee."),
             ("1–2 working days", "Domain & hosting", "Guided purchase, registered in your name."),
             ("2–3 working days", "Deployment", "Go-live with SSL, plus your mailbox on phone and laptop."),
             ("Launch day", "Handover", "Final review together and handover of all access."),
             (f"{cfg['project']['launch_warranty_days']} days", "Launch support", "Any reported issues fixed at no cost.")]
    y_line = Inches(3.35)
    rect(s, MX + Inches(0.15), y_line, W - 2 * MX - Inches(0.3), Pt(1.5), fill=LINE)
    cw = (W - 2 * MX) / 5
    for i, (when, h, d) in enumerate(steps):
        x = MX + cw * i
        last = i == len(steps) - 1
        rect(s, x + Inches(0.05), y_line - Inches(0.14), Inches(0.3), Inches(0.3),
             fill=ACCENT if not last else WHITE, line=ACCENT, shape=MSO_SHAPE.OVAL)
        text(s, x, Inches(2.75), cw, Inches(0.3), when.upper(), size=10, bold=True, color=ACCENT)
        text(s, x, Inches(3.75), cw - Inches(0.25), Inches(0.5), h, size=17, font=SERIF, color=INK)
        text(s, x, Inches(4.5), cw - Inches(0.3), Inches(1.0), d, size=11.5, color=MUTED, line_spacing=1.3)
    card(s, MX, Inches(5.55), W - 2 * MX, Inches(0.85), fill=BLUSH)
    text(s, MX + Inches(0.35), Inches(5.78), W - 2 * MX - Inches(0.7), Inches(0.5),
         [[("Launch quality checklist:  ", {"bold": True, "color": INK}),
           ("HTTPS everywhere  ·  mobile & browser tested  ·  speed reviewed  ·  SPF/DKIM/DMARC email  ·  forms verified  ·  backup handed over", {})]],
         size=12, color=TEXT)


def s_ownership(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, fill=DEEP)
    logo(s, MX, Inches(0.42), dark=True, size=Inches(0.42))
    eyebrow(s, MX, Inches(1.25), "06 · Our ownership promise", color=LILAC)
    text(s, MX, Inches(1.6), Inches(11.5), Inches(1.6),
         [[("Everything that runs your business online ", {}), ("belongs to you.", {"italic": True, "color": LILAC})]],
         size=30, font=SERIF, color=WHITE)
    accent(s, MX, Inches(2.65), color=LILAC)
    items = [("Domain", "Registered in your name with your contact details."),
             ("Hosting", "Account and billing in your name, paid to the provider."),
             ("Email", "Mailboxes on your domain, with passwords that are yours."),
             ("Access", "Delegated for setup only. Revoke it at any time.")]
    cw = (W - 2 * MX) / 4
    for i, (h, d) in enumerate(items):
        x = MX + cw * i
        rect(s, x, Inches(3.1), Pt(1.5), Inches(1.35), fill=ACCENT)
        text(s, x + Inches(0.25), Inches(3.1), cw - Inches(0.5), Inches(0.5), h, size=19, font=SERIF, color=WHITE)
        text(s, x + Inches(0.25), Inches(3.6), cw - Inches(0.5), Inches(0.9), d, size=11.5, color=RGBColor(0xC8, 0xC3, 0xDA), line_spacing=1.3)
    rect(s, MX, Inches(4.95), W - 2 * MX, Pt(0.6), fill=RGBColor(0x3A, 0x32, 0x6A))
    text(s, MX, Inches(5.15), Inches(4), Inches(0.3), "APPROX. THIRD-PARTY COSTS  ·  PAID DIRECTLY BY YOU", size=9, bold=True, color=LILAC)
    for i, tp in enumerate(cfg["pricing"]["third_party_estimates"]):
        x = MX + Inches(3.9) * i
        text(s, x, Inches(5.5), Inches(3.6), Inches(0.6),
             [[("approx. ", {"size": 12, "font": SANS, "color": LILAC}), (tp["estimate"].replace("approx. ", ""), {})]],
             size=26, font=SERIF, color=WHITE)
        text(s, x, Inches(6.1), Inches(3.6), Inches(0.3), tp["item"], size=11, color=RGBColor(0xC8, 0xC3, 0xDA))
    text(s, MX + Inches(7.8), Inches(5.5), Inches(3.9), Inches(1),
         [[("0% commission.", {"bold": True, "color": WHITE})], "Mailbox subscriptions are paid to the email provider as per their pricing."],
         size=12, color=RGBColor(0xC8, 0xC3, 0xDA), line_spacing=1.3)
    text(s, W - MX - Inches(2), H - Inches(0.5), Inches(2), Inches(0.3), f"{n:02d} / {t:02d}", size=8.5,
         color=RGBColor(0x8A, 0x84, 0xA6), align=PP_ALIGN.RIGHT)


def s_investment(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Your investment")
    pr = cfg["pricing"]
    cur = pr["currency"]
    title(s, "07 · Investment", "Simple, ", italic_tail="transparent pricing.")
    text(s, MX, Inches(2.7), Inches(11.5), Inches(0.4),
         [[("As a referral from ", {}), (cfg["client"]["referred_by"], {"bold": True, "color": ACCENT}),
           (", you receive preferential pricing. Because the deployment charges were not communicated clearly at the start, we have waived the deployment fee in full.", {})]],
         size=13, color=MUTED, line_spacing=1.3)
    rows = [("Website development", f"{len(cfg['client']['site_pages'])}-page website, built from your content",
             pr["development"]["standard"], pr["development"]["offered"]),
            ("Deployment", "Waived  ·  includes 1 professional mailbox", pr["deployment"]["standard"], pr["deployment"]["offered"])]
    y0 = Inches(3.55)
    for j, (hdr, x, al) in enumerate([("SERVICE", MX, PP_ALIGN.LEFT), ("STANDARD", MX + Inches(3.5), PP_ALIGN.RIGHT),
                                       ("YOUR RATE", MX + Inches(5.4), PP_ALIGN.RIGHT)]):
        text(s, x, y0, Inches(1.6) if j else Inches(4), Inches(0.3), hdr, size=9, bold=True, color=MUTED, align=al)
    tw = Inches(7.0)
    rect(s, MX, y0 + Inches(0.32), tw, Pt(1.2), fill=INK)
    for i, (a, b, std, off) in enumerate(rows):
        y = y0 + Inches(0.5) + Inches(0.78) * i
        text(s, MX, y, Inches(4.4), Inches(0.4), a, size=16, color=INK, font=SERIF)
        text(s, MX, y + Inches(0.36), Inches(4.4), Inches(0.3), b, size=10.5, color=ACCENT if i else MUTED)
        st = text(s, MX + Inches(3.5), y + Inches(0.05), Inches(1.6), Inches(0.4), money(std, cur), size=13,
                  color=RGBColor(0xA2, 0x9C, 0x9E), align=PP_ALIGN.RIGHT)
        st.text_frame.paragraphs[0].runs[0].font._rPr.set("strike", "sngStrike")
        text(s, MX + Inches(5.4), y + Inches(0.02), Inches(1.6), Inches(0.4), money(off, cur), size=18,
             bold=True, color=INK, align=PP_ALIGN.RIGHT)
        rect(s, MX, y + Inches(0.7), tw, Pt(0.6), fill=LINE)
    yt = y0 + Inches(2.1)
    text(s, MX, yt, Inches(4), Inches(0.4), "Total, payable on signing", size=13, bold=True, color=INK)
    text(s, MX + Inches(4.4), yt - Inches(0.05), Inches(2.6), Inches(0.5),
         money(pr["development"]["offered"] + pr["deployment"]["offered"], cur), size=22, font=SERIF, color=ACCENT, align=PP_ALIGN.RIGHT)
    # right: optional extra mailboxes
    x = MX + Inches(7.5)
    card(s, x, Inches(3.55), Inches(4.23), Inches(2.9), fill=BLUSH)
    eyebrow(s, x + Inches(0.35), Inches(3.8), "Optional", w=Inches(3))
    text(s, x + Inches(0.35), Inches(4.08), Inches(3.7), Inches(0.5), "Additional professional emails", size=16, font=SERIF, color=INK)
    text(s, x + Inches(0.35), Inches(4.55), Inches(3.7), Inches(0.6),
         [[(money(pr["extra_mailbox_fee"], cur), {}), ("  per account", {"size": 11, "font": SANS, "color": MUTED})]],
         size=24, font=SERIF, color=ACCENT)
    text(s, x + Inches(0.35), Inches(5.2), Inches(3.6), Inches(1.1),
         "One-time setup fee paid to OBD. The mailbox subscription is paid by you directly to the email provider.",
         size=11, color=MUTED, line_spacing=1.3)

def s_support(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Support")
    title(s, "08 · After launch", "Support you ", italic_tail="can rely on.")
    card(s, MX, Inches(2.75), Inches(4.3), Inches(3.75), fill=ACCENT, border=ACCENT)
    text(s, MX + Inches(0.4), Inches(3.05), Inches(3.6), Inches(0.3), "LAUNCH SUPPORT", size=9.5, bold=True, color=LILAC)
    text(s, MX + Inches(0.4), Inches(3.4), Inches(3.6), Inches(1.0), f"{cfg['project']['launch_warranty_days']} days", size=40, font=SERIF, color=WHITE)
    text(s, MX + Inches(0.4), Inches(4.5), Inches(3.5), Inches(1.8),
         [[("Any issue you report after go-live is fixed at no cost.", {"bold": True, "color": WHITE})],
          "Further changes and ongoing maintenance are available on request, quoted in advance."],
         size=11.5, color=LILAC, line_spacing=1.3, spacing=6)
    x = MX + Inches(4.65)
    rows = [("Critical", "Website or email down", "Same working day"),
            ("High", "Key page or form broken", "Within 1 working day"),
            ("Normal", "Question or request", "Within 1 working day")]
    text(s, x, Inches(2.75), Inches(2), Inches(0.3), "PRIORITY", size=9, bold=True, color=MUTED)
    text(s, x + Inches(1.6), Inches(2.75), Inches(3), Inches(0.3), "EXAMPLE", size=9, bold=True, color=MUTED)
    text(s, x + Inches(4.9), Inches(2.75), Inches(2.2), Inches(0.3), "FIRST RESPONSE", size=9, bold=True, color=MUTED, align=PP_ALIGN.RIGHT)
    rect(s, x, Inches(3.07), Inches(7.08), Pt(1.2), fill=INK)
    for i, (a, b, c_) in enumerate(rows):
        y = Inches(3.25) + Inches(0.6) * i
        text(s, x, y, Inches(1.5), Inches(0.4), a, size=14, bold=True, color=ACCENT if i == 0 else INK)
        text(s, x + Inches(1.6), y, Inches(3.3), Inches(0.4), b, size=13, color=TEXT)
        text(s, x + Inches(4.9), y, Inches(2.18), Inches(0.4), c_, size=13, color=INK, align=PP_ALIGN.RIGHT)
        rect(s, x, y + Inches(0.45), Inches(7.08), Pt(0.6), fill=LINE)
    text(s, x, Inches(5.2), Inches(7.08), Inches(1.3),
         [[("Contact:  ", {"bold": True, "color": INK}), (f"WhatsApp project group  ·  {cfg['project']['account_manager']} (Account Manager)  ·  {cfg['company']['founder']} (escalations)", {})],
          [("Hours:  ", {"bold": True, "color": INK}), ("Monday to Saturday, 10:00–19:00 PKT", {})]],
         size=12, color=MUTED, line_spacing=1.3, spacing=6)

def s_next(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    chrome(s, cfg, n, t, "Next steps")
    title(s, "09 · Next steps", "Three steps ", italic_tail="to launch.")
    pr = cfg["pricing"]
    steps = [("Sign", "Review and sign the Service Agreement."),
             ("Pay", f"Settle the development fee of {money(pr['development']['offered'], pr['currency'])}."),
             ("Launch", "Short call to purchase the domain & hosting. We then take the website live.")]
    cw = (W - 2 * MX - Inches(0.3)) / 3
    for i, (h, d) in enumerate(steps):
        x = MX + (cw + Inches(0.15)) * i
        card(s, x, Inches(2.75), cw, Inches(2.25), fill=BLUSH)
        text(s, x + Inches(0.35), Inches(2.95), Inches(1), Inches(0.8), str(i + 1), size=38, font=SERIF, color=ACCENT)
        text(s, x + Inches(0.35), Inches(3.8), cw - Inches(0.6), Inches(0.4), h, size=18, font=SERIF, color=INK)
        text(s, x + Inches(0.35), Inches(4.25), cw - Inches(0.6), Inches(0.8), d, size=12, color=MUTED, line_spacing=1.25)
    text(s, MX, Inches(5.35), Inches(4), Inches(0.3), "YOUR DOCUMENTS", size=9, bold=True, color=ACCENT)
    docs = ["Welcome Pack", "Service Agreement", "Go-Live & Handover"]
    runs = []
    for i, d in enumerate(docs):
        runs += [(f"0{i+1} ", {"bold": True, "color": ACCENT}), (d + ("      " if i < len(docs) - 1 else ""), {})]
    text(s, MX, Inches(5.7), W - 2 * MX, Inches(0.4), [runs], size=11, color=INK)

def s_thanks(prs, cfg, n, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, fill=DEEP)
    picture(s, ROOT / "assets" / "logo" / "circuit-strip.jpg", W - Inches(3.6), 0, Inches(3.6), H)
    logo(s, MX, Inches(0.7), dark=True)
    cl = cfg["client"]
    text(s, MX, Inches(2.3), Inches(8.8), Inches(1.2), [[("Thank you, ", {}), (f"{cl['business']}.", {"italic": True, "color": LILAC})]],
         size=50, font=SERIF, color=WHITE)
    accent(s, MX, Inches(3.65), color=LILAC)
    text(s, MX, Inches(3.95), Inches(8.5), Inches(1),
         f"{cl['tagline']} Let's give that promise a home online.",
         size=17, color=RGBColor(0xDD, 0xD3, 0xF7))
    co = cfg["company"]
    rect(s, MX, H - Inches(1.45), Inches(8.3), Pt(0.6), fill=RGBColor(0x3A, 0x32, 0x6A))
    text(s, MX, H - Inches(1.2), Inches(6), Inches(0.4), co["founder"], size=18, font=SERIF, color=WHITE)
    text(s, MX, H - Inches(0.8), Inches(8.3), Inches(0.4), f"{co['founder_title']}  ·  {co['founder_secondary']}", size=10.5, color=LILAC)
    contact = "  ·  ".join(v for v in (co["email"], co["phone"], co["website"]) if v)
    if contact:
        text(s, MX, H - Inches(0.5), Inches(8.3), Inches(0.3), contact, size=10, color=WHITE)


def build(cfg, out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    body = [s_about, s_experience, s_team, s_status, s_roadmap, s_ownership, s_investment,
            s_support, s_next]
    total = len(body) + 2
    s_title(prs, cfg)
    for i, fn in enumerate(body, start=2):
        fn(prs, cfg, i, total)
    s_thanks(prs, cfg, total, total)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path
