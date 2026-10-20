#!/usr/bin/env python3
"""New decks and templates for Decks.

    python3 newdeck.py templates         list the templates
    python3 newdeck.py build <template>  print a template's assembled deck file (for checking)

The engine (templates/_engine/engine.html) is one self-contained deck file without
content: the layouts, editing, versions, sorter, present mode, and "Save a copy" are
there; the slides, figures, and versions are placeholders. Each template in
templates/<id>/ fills them:

    template.json   name, description, kind, the sidebar line, and which files to copy
    content.js      CATS, CAT_LABELS, CAT_SHORT, FIGS, LIBRARY (the slides)
    pitches.js      PITCHES (the versions)
    anything else   copied next to the deck (img/, slide data scripts)

create_deck() copies a template into a new folder in the decks folder (sbconfig.py),
gives it its own storage prefix, adds it to decks.json, starts its git history, and
writes BRIEF.md with the brief from the New deck form. The server and `sb new` both
call it. Every new deck also carries the basic guided layouts in _engine/basics.js.
"""
import html
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

import sbconfig

APP = Path(__file__).resolve().parent
TEMPLATES = APP / 'templates'
ENGINE = TEMPLATES / '_engine' / 'engine.html'   # the bundled engine every deck is built from
_REG = threading.Lock()          # decks.json: read-check-reserve and write are short; building the deck happens outside
_RESERVED = set()                # prefixes handed out to decks still being built
GITIGNORE = '*.pdf\n*.zip\n*.pptx\n_archive/\n.DS_Store\n.slidebuilder/thumb.png\n'
IMPORT = '_import'               # the hidden template Import a deck starts from
IMPORT_TYPES = {'.pptx': 'PowerPoint', '.pdf': 'PDF'}
IMPORT_NOT = {   # files Import a deck can't read, and how to get one it can
    '.key': 'That’s a Keynote file. In Keynote, choose File › Export To › PowerPoint, then import the .pptx file.',
    '.ppt': 'That’s an older PowerPoint file (.ppt). Open it in PowerPoint, save it as .pptx (File › Save As), then import that.',
    '.pps': 'That’s a PowerPoint show (.pps). Open it in PowerPoint, save it as .pptx, then import that.',
    '.ppsx': 'That’s a PowerPoint show (.ppsx). Open it in PowerPoint, save it as .pptx, then import that.',
    '.odp': 'That’s an OpenDocument deck. Save it as PowerPoint (.pptx) in LibreOffice or Keynote, then import that.',
}

# ---------- themes ----------
# Each theme sets the engine's colors (CSS variables, written into the deck as one <style> block) and
# a color map for converted PowerPoint slides: their Crimson palette mapped onto the theme, text
# and fills separately, so Dark can turn dark text light and white panels dark.
def _pp(accent, deep, tint, light, dark, gray, mid, muted, panel, line, bg='#FFFFFF', card='#FFFFFF', on_dark='#FFFFFF', text_dark=None, text_gray=None, cover=None):
    fill = {'#A22731': accent, '#A12731': accent, '#842F2F': deep, '#F6E9EA': tint, '#F3D9DC': tint, '#E8DDDD': tint,
            '#D98A92': light, '#E07A84': light, '#272323': dark, '#272222': dark, '#C9C4C4': gray, '#F4F2F2': panel,
            '#E4E1E1': line, '#FFFFFF': card, '#5A5656': mid, '#8A8585': muted}
    text = dict(fill, **{'#FFFFFF': on_dark, '#272323': text_dark or dark, '#272222': text_dark or dark, '#C9C4C4': text_gray or gray})
    return {'fill': fill, 'text': text, 'bg': {'#FFFFFF': bg, '#272323': cover or dark, '#A12731': accent, '#A22731': accent}}

THEMES = {
    'navy': {'name': 'Navy', 'accent': '#2E6DA4',
             'vars': {'--text-mute': '#66707D', '--char': '#1B2A3E', '--char-2': '#24364F', '--maroon': '#1F3A5F', '--maroon-deep': '#142740', '--red': '#3B82C4', '--red-deep': '#2E6DA4',
                      '--paper': '#F4F6F9', '--light': '#EEF1F5', '--title-red': '#1F3A5F', '--accent': '#2E6DA4', '--lime': '#3B82C4', '--line': '#D3DAE3', '--text-light': '#5B6675'},
             'div': {'--div-accent': '#8DB5DD'},
             'pp': _pp('#2E6DA4', '#1F3A5F', '#E3ECF6', '#8DB5DD', '#1B2A3E', '#C6CFDA', '#4F5B6B', '#7D8896', '#F1F4F8', '#DDE3EA')},
    # LBMC: the LBMC Tahoe palette (LBMC Blue #009ADD, cool grays) in the look of the Digital Steering Committee deck:
    # charcoal #1E2831 for dark slides, light titles under a small blue eyebrow, grey panels, a quiet footer, no logo on
    # content slides. Instrument Sans by the user's choice (Tahoe says Geist). Text blue is #0078AD: #009ADD is too
    # light for small text on white, so it is kept for fills and for text on charcoal.
    'lbmc': {'name': 'LBMC', 'accent': '#009ADD', 'footer': True,
             'fonts': 'https://fonts.googleapis.com/css2?family=Instrument+Sans:ital,wght@0,400..700;1,400..700&display=swap',
             'vars': {'--text-mute': '#6A7785', '--char': '#1E2831', '--char-2': '#3A4451', '--maroon': '#1E2831', '--maroon-deep': '#141B22', '--red': '#009ADD', '--red-deep': '#0078AD',
                      '--paper': '#F6F8FA', '--light': '#FFFFFF', '--ink': '#1F2832', '--text': '#1F2832', '--title-red': '#1F2832', '--accent': '#009ADD', '--lime': '#009ADD', '--line': '#DCE3EA', '--text-light': '#4A5663',
                      '--sans': "'Instrument Sans', -apple-system, BlinkMacSystemFont, sans-serif", '--slide': "'Instrument Sans', -apple-system, sans-serif",
                      '--foot-text': '"LBMC"'},
             'div': {'--div-bg': '#1E2831', '--div-accent': '#009ADD'},
             'css': [
                 # type: titles light (400) and one size on every content slide; body regular; bold is semibold.
                 # Four sizes in all (.75rem labels, 1rem body, 1.6rem titles, big type on cover and closing) to pass Checks
                 '.slide .h-title { font-weight: 400; letter-spacing: -0.005em; font-style: normal; }',
                 '.slide:not(.l-title):not(.l-section):not(.l-closing) .h-title { font-size: 1.6rem; line-height: 1.25; }',
                 '.slide b, .slide strong { font-weight: 600; }',
                 '.slide .eyebrow { font: 600 .75rem/1.2 var(--slide); letter-spacing: .14em; text-transform: uppercase; color: var(--red-deep); margin: 0 0 .5rem; }',
                 # footer: the deck name on the left and the page number on the right, small and grey, no rule.
                 # Cover, section, and closing slides go without. The logo is on the cover only
                 '#deck { counter-reset: lbmc-page; } #deck > .slide-wrap { counter-increment: lbmc-page; }',
                 '.slide:not(.l-title) > .logo { display: none; }',
                 '.slide:not(.l-title):not(.l-closing):not(.l-section)::before { content: var(--foot-text); position: absolute; left: 44px; bottom: 16px; font: 400 10px/1 var(--slide); color: var(--text-mute); text-align: left; pointer-events: none; z-index: 9; }',
                 '.slide:not(.l-title):not(.l-closing):not(.l-section)::after { content: counter(lbmc-page); position: absolute; right: 44px; bottom: 16px; font: 400 10px/1 var(--slide); color: var(--text-mute); font-variant-numeric: tabular-nums; pointer-events: none; z-index: 9; }',
                 '.slide.dark::before, .slide.maroon::before, .slide.l-reflect::before, .slide.dark::after, .slide.maroon::after, .slide.l-reflect::after { color: #8B95A1; }',
                 # cover: charcoal, blue eyebrow, a light white title, the line under it, then labelled facts and the presenter
                 '.slide.l-title { background: var(--char) !important; color: #fff; padding: 3rem 3.4rem; justify-content: center; }',
                 '.l-title .kicker { font: 600 .75rem/1.2 var(--slide); letter-spacing: .16em; text-transform: uppercase; color: var(--red); }',
                 '.l-title .h-title { color: #fff; font-size: 2.8rem; max-width: none; margin-top: .7rem; }',
                 '.l-title .date { margin-top: 1.1rem; font-size: 1rem; font-weight: 400; line-height: 1.5; color: #C7D0D9; max-width: 62%; }',
                 '.l-title .meta { display: grid; grid-template-columns: repeat(3, max-content); column-gap: 3.2rem; margin-top: 1.8rem; padding-top: 1.3rem; border-top: 1px solid rgba(255,255,255,.16); width: fit-content; }',
                 '.l-title .meta small { display: block; font: 600 .75rem/1.2 var(--slide); letter-spacing: .14em; text-transform: uppercase; color: #8B95A1; margin-bottom: .45rem; }',
                 '.l-title .meta span { display: block; font-size: 1rem; color: #fff; }',
                 '.l-title .by { position: absolute; left: 3.4rem; bottom: 2.2rem; font-size: .75rem; color: #8B95A1; }',
                 '.l-title .logo-plate { right: 3.4rem; bottom: 2rem; } .l-title .logo { width: 72px; }',
                 # closing: the cover's charcoal, left-aligned, a short blue rule, then name and role
                 '.slide.l-closing { background: var(--char) !important; align-items: flex-start; justify-content: center; text-align: left; padding: 3rem 3.4rem; }',
                 '.l-closing .h-title { color: #fff; font-size: 2.8rem; margin: 0 0 1.4rem; }',
                 ".l-closing .h-title::after { content: ''; display: block; width: 3.5rem; height: 2px; background: var(--red); margin-top: 1.2rem; }",
                 '.l-closing .contacts { flex-direction: column; align-items: flex-start; gap: .35rem; font-size: 1rem; color: #C7D0D9; }',
                 '.l-closing .contacts span:first-of-type { color: #fff; font-weight: 600; } .l-closing .contacts .sep { display: none; }',
                 '.l-closing .site { position: absolute; left: 3.4rem; bottom: 2.2rem; margin: 0; font-size: .75rem; color: #8B95A1; }',
                 # section dividers and question slides: charcoal, left-aligned, light type
                 '.slide.l-section .h-title { font-weight: 400; }',
                 '.slide.l-reflect { align-items: flex-start; text-align: left; padding: 3rem 4rem; }',
                 '.l-reflect .q { font-weight: 400; font-size: 2.2rem; max-width: 30ch; } .l-reflect .how { font-weight: 600; font-size: .75rem; letter-spacing: .14em; color: var(--red); }',
                 # two columns: top-aligned, small blue uppercase headings, 1rem body; cards become grey panels with a hairline
                 '.l-twocol .cols { align-content: start; } .l-twocol .cols.one { grid-template-columns: minmax(0, 620px); justify-content: start; }',
                 '.l-twocol .col h4, .l-twocol .cols.bigcols .col h4 { font-size: .75rem; font-weight: 600; letter-spacing: .12em; text-transform: uppercase; color: var(--red-deep); margin: 0 0 .7rem; }',
                 '.l-twocol .body-p, .l-twocol .cols.bigcols .body-p, .l-data .body-p { font-size: 1rem; line-height: 1.5; margin-bottom: .55rem; }',
                 '.l-plain .body-p { font-size: 1rem; line-height: 1.55; } .l-plain .content { justify-content: flex-start; padding-top: .6rem; }',
                 '.l-twocol .col.card, .l-twocol .cols.bigcols .col.card, .l-flow .step, .l-ddiag .dd-box, .l-transform .card, .l-ladder .facts, .l-card .panel, .l-breakout .card, .l-dialogue .bub { background: var(--paper); box-shadow: none !important; border: 1px solid var(--line); border-radius: 2px; padding: 1.1rem 1.3rem; }',
                 # tables: a slim charcoal header, hairline rows, regular cells, no shadow, rounding, or zebra stripes
                 '.l-data table, .l-pptx table.pp-table { box-shadow: none; border-radius: 0; background: transparent; font-size: 1rem; }',
                 '.l-data table { flex: none; }',
                 '.l-data th { background: var(--char); font-weight: 600; font-size: .75rem; text-transform: uppercase; letter-spacing: .1em; padding: .55rem 1rem; }',
                 '.l-data td { font-weight: 400; padding: .6rem 1rem; border-bottom: 1px solid var(--line); }',
                 '.l-data td:first-child, .l-data td.n, .l-data td.em { font-weight: 600; font-style: normal; }',
                 '.l-data tr:nth-child(even) td, .l-pptx table.pp-table tr:nth-child(even) td { background: transparent; }',
                 '.l-data tr:last-child td { border-bottom: 1px solid var(--line); }',
                 '.l-data .foot { font-size: .75rem; font-style: normal; color: var(--text-light); }',
                 # charts: the caption is a small grey label, as in the LBMC decks
                 '.l-chart .cap { font: 600 .75rem/1.2 var(--slide); letter-spacing: .1em; text-transform: uppercase; color: var(--text-light); font-style: normal; margin: .2rem 0 .4rem; }'],
             'pp': _pp('#009ADD', '#1E2831', '#E6F5FB', '#99D7F1', '#1E2831', '#C7D0D9', '#4A5663', '#6A7785', '#F6F8FA', '#DCE3EA')},
    'dark': {'name': 'Dark', 'accent': '#E0505A', 'dark': True,
             'vars': {'--char': '#141211', '--char-2': '#2B2726', '--maroon': '#2B2726', '--maroon-deep': '#141211', '--red': '#EC6670', '--red-deep': '#E0505A', '--accent': '#E0505A', '--lime': '#EC6670'},
             'slide': {'--paper': '#262322', '--light': '#1E1C1B', '--white': '#262322', '--ink': '#F3F0EE', '--text': '#E6E1DE', '--text-light': '#B5ADA9', '--text-mute': '#A39C98',
                       '--line': '#3A3634', '--title-red': '#FFFFFF'},
             'div': {'--div-bg': '#141211', '--div-accent': '#E0505A'},
             'pp': _pp('#E0505A', '#B23A43', '#3D2427', '#A8474F', '#3A3533', '#57514F', '#C9C2BE', '#9A938F', '#2B2826', '#3A3634',
                       bg='#1E1C1B', card='#262322', on_dark='#FFFFFF', text_dark='#F3F0EE', text_gray='#B5ADA9', cover='#141211')},
    'minimal': {'name': 'Minimal', 'accent': '#C8102E',
                'vars': {'--text-mute': '#6B6B6B', '--char': '#111111', '--char-2': '#222222', '--maroon': '#111111', '--maroon-deep': '#000000', '--red': '#E0293F', '--red-deep': '#C8102E',
                         '--paper': '#F5F5F5', '--light': '#FFFFFF', '--title-red': '#111111', '--accent': '#C8102E', '--lime': '#E0293F', '--line': '#E6E6E6', '--text-light': '#555555'},
                'div': {'--div-bg': '#FFFFFF', '--div-fg': '#111111', '--div-accent': '#C8102E'},
                'pp': _pp('#C8102E', '#8E0B21', '#FBE7EA', '#E8808F', '#111111', '#D6D6D6', '#555555', '#888888', '#F5F5F5', '#E6E6E6')},
    'forest': {'name': 'Forest', 'accent': '#2F7A55',
               'vars': {'--text-mute': '#64706A', '--char': '#16211C', '--char-2': '#1E2C25', '--maroon': '#1E4D3A', '--maroon-deep': '#133326', '--red': '#3C9A6B', '--red-deep': '#2F7A55',
                        '--paper': '#F5F8F6', '--light': '#EEF2EF', '--title-red': '#1E4D3A', '--accent': '#2F7A55', '--lime': '#3C9A6B', '--line': '#D2DCD5', '--text-light': '#56645C'},
               'div': {'--div-accent': '#9FD4B8'},
               'pp': _pp('#2F7A55', '#1E4D3A', '#E2F0E8', '#8CC3A6', '#1C2B24', '#C5CEC8', '#4F5C55', '#7E8A83', '#F1F5F2', '#DCE4DF')},
    'teal': {'name': 'Teal', 'accent': '#13838F',
             'vars': {'--text-mute': '#62767A', '--char': '#102227', '--char-2': '#173038', '--maroon': '#0F4C5C', '--maroon-deep': '#0A3540', '--red': '#1FA2AE', '--red-deep': '#13838F',
                      '--paper': '#F4F8F9', '--light': '#EDF3F4', '--title-red': '#0F4C5C', '--accent': '#13838F', '--lime': '#1FA2AE', '--line': '#CFDDE0', '--text-light': '#4C5E62'},
             'div': {'--div-accent': '#8FD6DC'},
             'pp': _pp('#13838F', '#0F4C5C', '#DDF0F2', '#7CC3CA', '#12262B', '#C3D0D2', '#4C5E62', '#7B8B8E', '#EFF5F6', '#D8E3E5')},
    'plum': {'name': 'Plum', 'accent': '#7B3F72',
             'vars': {'--text-mute': '#75697A', '--char': '#221520', '--char-2': '#2E1D2B', '--maroon': '#4A2545', '--maroon-deep': '#32182F', '--red': '#9C5292', '--red-deep': '#7B3F72',
                      '--paper': '#F8F5F7', '--light': '#F3EFF2', '--title-red': '#4A2545', '--accent': '#7B3F72', '--lime': '#9C5292', '--line': '#DDD2DA', '--text-light': '#5E5260'},
             'div': {'--div-accent': '#E2B8DA'},
             'pp': _pp('#7B3F72', '#4A2545', '#F1E6EF', '#C08DB8', '#2A1C28', '#CFC6CD', '#5E5260', '#8C808A', '#F6F2F5', '#E4DCE2')},
}
SERIF = "Georgia, 'Times New Roman', serif"


def theme_key(style):
    s = (style or '').strip().lower()
    return next((k for k, t in THEMES.items() if s in (k, t['name'].lower())), None)


def theme_previews():
    """The colors the New deck form draws each theme's sample slide with."""
    out = []
    for key, t in THEMES.items():
        v, sl = t['vars'], t.get('slide', {})
        out.append({'key': key, 'name': t['name'], 'dark': bool(t.get('dark')), 'bg': sl.get('--light', '#FFFFFF'),
                    'title': sl.get('--title-red') or v.get('--title-red', '#272323'), 'text': sl.get('--text-light') or v.get('--text-light', '#5A5656'),
                    'accent': t['accent'], 'block': v.get('--maroon', '#842F2F'), 'line': sl.get('--line') or v.get('--line', '#E4E1E1')})
    return out


def theme_css(key, serif=False, originals=False, foot=None):
    """originals: an imported deck. Its converted slides keep the "original" button (a picture of each slide as it
    was) and their own logos, so the deck's logo stays off them."""
    t = THEMES.get(key)
    css = (['.slide.l-pptx > .logo, .slide.l-image > .logo { display: none; }   /* imported slides carry their own logos */'] if originals
           else ['.ctrls .cbtn.orig { display: none; }   /* the PowerPoint picture behind a converted slide: not useful in a new deck */'])
    if not t and not serif:
        return '<style id="deck-theme">\n' + css[0] + '\n</style>'
    css += [f'/* Deck theme: {t["name"] if t else "engine colors"}{", serif titles" if serif else ""}. Change these to restyle the whole deck. */']
    if t:
        if t.get('fonts'):   # the theme's typeface: @import has to come before every other rule in the block
            css.insert(0, f"@import url('{t['fonts']}');")
        css.append(':root { ' + ' '.join(f'{k}: {v};' for k, v in t['vars'].items()) + ' }')
        if t.get('slide'):   # text and background colors that belong to the slides only, not the editor around them
            css.append('.slide, .slide-frame { ' + ' '.join(f'{k}: {v};' for k, v in t['slide'].items()) + ' }')
        css.append('.slide.l-section { ' + ' '.join(f'{k}: {v};' for k, v in t.get('div', {}).items()) + ' }')
        bg = t.get('div', {}).get('--div-bg', '')
        if re.fullmatch(r'#[0-9A-Fa-f]{6}', bg) and sum(int(bg[i:i + 2], 16) for i in (1, 3, 5)) > 600:
            css.append('.slide.l-section .logo { filter: none; opacity: 1; }   /* a light divider keeps the logo in its own colors */')
        # a few layouts use fixed colors; bring them onto the theme
        css.append(f'.slide .step.accent, .slide .flow-body .step.accent {{ background: var(--maroon) !important; border-color: var(--maroon) !important; }}')
        css.append('.slide table th { background: var(--maroon) !important; }')
        css.append('.slide.maroon, .slide.l-reflect:not(.dark), .slide.l-breakout { background: var(--maroon) !important; }')
        if t.get('dark'):
            css.append('.slide { background: var(--light); color: var(--text); }')
            css.append('.slide .person, .slide .card, .slide .step:not(.accent), .slide table td, .slide table tr { background: #262322 !important; color: var(--text) !important; }')
            css.append('.slide .logo { filter: brightness(0) invert(1); opacity: .8; }')
        css += t.get('css', [])
        if t.get('footer') and foot:   # the footer line on every content slide: the deck's name
            css.append(':root { --foot-text: ' + json.dumps(str(foot), ensure_ascii=False).replace('</', '<\\/') + '; }')
    if serif:
        css.append(f'.slide .h-title, .slide h1, .slide h2 {{ font-family: {SERIF}; font-weight: 700; letter-spacing: 0; }}')
    return '<style id="deck-theme">\n' + '\n'.join(css) + '\n</style>'


def theme_js(key, serif=False):
    t = THEMES.get(key)
    return 'const DECK_THEME = ' + json.dumps({'pp': t['pp'] if t else None, 'serif': SERIF if serif else None}) + ';'


def recolor_images(folder, key):
    """Converted chart pictures: the Crimson reds move to the theme's accent (Dark keeps its light panels)."""
    t = THEMES.get(key)
    if not t:
        return 0
    try:
        import numpy as np
        from PIL import Image
    except ImportError:
        return 0
    acc = np.array([int(t['accent'][i:i + 2], 16) for i in (1, 3, 5)], dtype=float)
    n = 0
    for f in sorted((folder / 'img').glob('ct-*')):
        if f.suffix.lower() not in ('.jpg', '.jpeg', '.png') or re.match(r'ct-(sp-)?\d{3}\.jpg$', f.name):
            continue   # the full-slide PowerPoint pictures: never shown in a new deck
        im = Image.open(f); mode = im.mode
        a = np.asarray(im.convert('RGBA'), dtype=float)
        rgb = a[..., :3] / 255
        mx, mn = rgb.max(-1), rgb.min(-1)
        sat = np.where(mx > 0, (mx - mn) / np.maximum(mx, 1e-6), 0)
        r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
        red = (r >= g) & (r >= b) & (r - np.maximum(g, b) > .03) & (np.abs(g - b) < .12)   # the Crimson reds and their pink tints
        if not red.any():
            continue
        amt = np.clip(sat / .76, 0, 1)[..., None]          # how much accent is in the pixel (Crimson #A22731 is full)
        shade = np.clip(mx / .64, 0, 1.3)[..., None]       # darker reds stay darker
        target = (1 - amt) * 255 + amt * np.clip(acc * shade, 0, 255)
        a[..., :3] = np.where(red[..., None], target, a[..., :3])
        out = Image.fromarray(a.astype('uint8'), 'RGBA')
        out = out.convert('RGB') if mode != 'RGBA' else out
        out.save(f, quality=88) if f.suffix.lower() in ('.jpg', '.jpeg') else out.save(f)
        n += 1
    return n


def js_str(v):
    """A JavaScript string literal that is also safe inside a <script> tag."""
    return json.dumps(str(v), ensure_ascii=False).replace('</', '<\\/').replace('\u2028', '\\u2028').replace('\u2029', '\\u2029')


# ---------- templates ----------

def list_templates():
    out = []
    for d in sorted(TEMPLATES.iterdir()) if TEMPLATES.exists() else []:
        m = d / 'template.json'
        if d.name.startswith('_') or not m.exists():
            continue
        t = json.loads(m.read_text(encoding='utf-8'))
        t['id'] = d.name
        out.append(t)
    return sorted(out, key=lambda t: t.get('order', 99))


def template(tid):
    for t in list_templates():
        if t['id'] == tid:
            return t
    if tid == IMPORT:   # hidden from the list: Import a deck uses it
        t = json.loads((TEMPLATES / IMPORT / 'template.json').read_text(encoding='utf-8'))
        t['id'] = IMPORT
        return t
    raise ValueError(f'No template called {tid!r}. Try one of: ' + ', '.join(t['id'] for t in list_templates()))


def pieces_js(pieces, template_pitches):
    """Versions for the pieces the user picked (the first is the main deck), then the template's own extra versions.
    A piece other than the main one starts as one placeholder slide (only:true, in:[piece]) until Claude builds it;
    a facilitator pack is the engine's facilitator version (the agenda plus fac:true slides)."""
    main, rest = pieces[0], pieces[1:]
    js = ['const TEMPLATE_PITCHES = (() => {', template_pitches, '  return PITCHES;', '})();',
          "const pieceOnly = (k) => Object.keys(LIBRARY).filter((id) => (LIBRARY[id].in || []).includes(k));"]
    for pc in rest:
        if pc['id'] == 'facilitator':
            continue
        title = f"{pc['name']}: {pc.get('desc', '').rstrip('.')}" if pc.get('desc') else pc['name']
        js.append(f"LIBRARY[{js_str(pc['id'] + '-start')}] = {{ cat: CATS[CATS.length - 1], layout: 'plain', only: true, in: [{js_str(pc['id'])}], "
                  f"title: {js_str(title)}, body: ['Claude builds this version from the main deck and the brief.', 'Pick it from the version menu at the top to see it.'], "
                  f"notes: 'Placeholder: Claude replaces this slide when it builds this version.' }};")
    v = [f"  {js_str(main['id'])}: {{ name: {js_str(main['name'])}, default: true, blocks: (Object.values(TEMPLATE_PITCHES).find((x) => x.default) || {{}}).blocks || CATS.flatMap(byCat) }},"]
    for pc in rest:
        blocks = 'facPack()' if pc['id'] == 'facilitator' else f"pieceOnly({js_str(pc['id'])})"
        v.append(f"  {js_str(pc['id'])}: {{ name: {js_str(pc['name'])}, blocks: {blocks} }},")
    ids = js_str([pc['id'] for pc in pieces])
    js += ['const PITCHES = Object.assign({', *v, '}, Object.fromEntries(Object.entries(TEMPLATE_PITCHES).filter(([k, x]) => !x.default && k !== \'facilitator\' && !' + ids + '.includes(k))));']
    return '\n'.join(js)


def assemble(tid, name, prefix, pieces=None, theme=None, serif=False, logo=None, originals=False):
    """A template's deck file, ready to open."""
    t = template(tid)
    d = TEMPLATES / tid
    engine = ENGINE.read_text(encoding='utf-8')
    shown = html.escape(name, quote=False)   # slide text is HTML: the name shows as typed, never as markup
    content = (d / 'content.js').read_text(encoding='utf-8').replace("'{{NAME_JS}}'", js_str(shown)).replace('{{NAME_JS}}', js_str(shown)[1:-1].replace("'", "\\'"))
    pitches = (d / 'pitches.js').read_text(encoding='utf-8')
    if pieces:
        pitches = pieces_js(pieces, pitches)
    scripts = '\n'.join(f'<script src="{html.escape(s)}"></script>' for s in t.get('scripts', []))
    basics = TEMPLATES / '_engine' / 'basics.js'
    if t.get('basics', True) and basics.exists():   # the basic guided layouts behind Add a slide
        content += '\n' + basics.read_text(encoding='utf-8')
    head = [theme_js(theme, serif)]
    if logo:
        head.append(f'const LOGO = {js_str(logo)}, LOGO_ALT = {js_str(name)};')
    content = '\n'.join(head) + '\n' + content
    out = (engine.replace('/*{{CONTENT}}*/', content).replace('/*{{PITCHES}}*/', pitches).replace('<!--{{THEME}}-->', theme_css(theme, serif, originals, foot=name))
           .replace('<!--{{SCRIPTS}}-->', scripts)
           .replace('{{TITLE}}', html.escape(name)).replace('{{BRAND}}', 'Strategy<span>U</span> · ' + html.escape(name))
           .replace('{{SUB}}', html.escape(t.get('sub', t['name']))).replace('{{PREFIX}}', prefix))
    assert '{{' not in out.replace('{{ ', ''), 'unfilled placeholder'
    return out


# ---------- new decks ----------

def slugify(name):
    s = re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')
    return s[:48].strip('-') or 'deck'


def new_prefix(name, taken):
    words = [w for w in re.findall(r'[a-z0-9]+', name.lower()) if w not in ('the', 'a', 'an', 'of', 'for', 'and', 'to', 'in')] or ['deck']
    base = ''.join(w[0] for w in words)[:3]
    base = (base + ''.join(words))[:3] if len(base) < 3 else base
    cand, n = base, 2
    while any(p.startswith(cand + '_') or (cand + '_').startswith(p) for p in taken):
        cand, n = f'{base}{n}', n + 1
    return cand + '_'


BRIEF_FIELDS = (('audience', 'Who it is for'), ('audienceType', 'Kind of audience'), ('knows', 'What they know and care about'),
                ('use', 'How it will be used'), ('length', 'Length'), ('point', 'The one point'), ('structure', 'Structure'),
                ('density', 'Density'), ('style', 'Style'), ('titles', 'Title font'), ('notes', 'Notes and source material'))


def tmpl_phrase(tmpl):
    n = tmpl['name']
    return f'the {n}' if 'template' in n.lower() else f'the {n} template'


def prompt_for(deck, brief, tmpl, pieces=None):
    bits = [f'New deck “{deck["name"]}” in Decks, from {tmpl_phrase(tmpl)}']
    if brief.get('length'):
        bits[0] += f', {brief["length"]}'
    if brief.get('audience'):
        bits[0] += f', for {brief["audience"]}'
    lines = [bits[0] + '.', f'It is set up at {deck["url"]} ({deck["folder"]}, slug {deck["slug"]}).']
    if brief.get('point'):
        lines.append(f'The one point: {brief["point"]}')
    shape = [brief.get('use'), brief.get('audienceType') and 'for ' + brief['audienceType'].lower(),
             brief.get('structure') and 'structure: ' + brief['structure'].lower(), brief.get('density') and brief['density'].lower() + ' density',
             brief.get('style') and 'style: ' + brief['style']]
    shape = '; '.join(s for s in shape if s)
    if shape:
        lines.append(f'Brief: {shape}.')
    if pieces and len(pieces) > 1:
        lines.append('Build these together, each as a version in the deck: ' + ', '.join(pc['name'] + (' (the main one, first)' if i == 0 else '') for i, pc in enumerate(pieces)) + '.')
    lines.append('Read BRIEF.md and shape the template to it (structure, density, style, and the slides themselves), then build a first draft. '
                 'Titles are plain sentences, two lines max, everything passes Checks › Guides, and mark any wording you drafted for me in the notes.')
    return ' '.join(lines)


def check_source(source):
    """Import a deck: the file to bring in, checked before anything is made. Returns (path, its name, mode)."""
    path = Path(os.path.expanduser(str(source.get('path') or '').strip().strip('"\''))) if source.get('path') else None
    if not path:
        raise ValueError('Choose the deck to import: a PowerPoint file (.pptx) or a PDF.')
    name = re.sub(r'[\\/:\x00-\x1f]', '-', str(source.get('name') or path.name)).strip(' .') or path.name
    ext = Path(name).suffix.lower()
    if ext in IMPORT_NOT:
        raise ValueError(IMPORT_NOT[ext])
    if ext not in IMPORT_TYPES:
        raise ValueError('Import a deck takes a PowerPoint file (.pptx) or a PDF. From Keynote or Google Slides, export it to PowerPoint first.')
    if not path.is_file():
        raise ValueError(f'There’s no file at {source.get("path")}. Check the path, or choose the file instead.')
    with open(path, 'rb') as f:
        head = f.read(5)
    if (ext == '.pptx' and head[:2] != b'PK') or (ext == '.pdf' and head[:4] != b'%PDF'):
        raise ValueError(f'{name} doesn’t look like a {IMPORT_TYPES[ext]} file. Export it again, then import that.')
    mode = 'rebuild' if source.get('mode') == 'rebuild' else 'keep'
    return path, name, mode


def size_label(n):
    return f'{n / 1e6:.0f} MB' if n >= 1e6 else f'{max(n, 1000) / 1e3:.0f} KB'


def import_brief(deck, src, mode, brief):
    """BRIEF.md for an imported deck: what the user asked for, and the steps that bring it in."""
    slug, app, kind = deck['slug'], sbconfig.tilde(APP), IMPORT_TYPES[Path(src).suffix.lower()]
    pdf = kind == 'PDF'
    look = [brief.get('style') and f'the {brief["style"]} color set', brief.get('titles') == 'Serif' and 'serif titles', deck.get('logo') and 'the logo in img/']
    look = ', '.join(x for x in look if x)
    md = [f'# {deck["name"]}', '', f'Imported {time.strftime("%B %-d, %Y")} from `{src}` ({kind}, {deck["sourceSize"]}).', '', '## What to do', '']
    if mode == 'keep':
        md += ['Bring every slide in and keep its look: each slide converted so every word on it can be edited' + ('' if pdf else ', with its speaker notes')
               + '. Don’t rewrite anything yet. The user reviews it in Decks first.', '']
    else:
        md += [f'Rebuild it in Decks’ own layouts{" and " + look if look else ""}: one slide for each slide in the file, in the same order, '
               'with the same words and the same speaker notes.', '']
    if brief.get('notes'):
        md += ['## Notes from the user', '', brief['notes'], '']
    md += ['## How', '', f'Run these from the Decks folder, `{app}`. The Decks skill has the same steps under “Importing a deck”.', '',
           '1. `python3 tools/importdeck.py --check`. If it says Python packages are missing, ask the user before installing them '
           '(`python3 -m pip install python-pptx lxml Pillow PyMuPDF`).']
    if mode == 'keep':
        if pdf:
            md.append('2. A PDF converts on its own: each page is drawn without its words, and the words go back on top as text, one line per text box.')
        else:
            md.append('2. A PDF of the same deck in `source/` makes it come out right: charts, SmartArt, and icons are cut from it, and each slide gets a '
                      'picture of the original to compare with. With none there, the tool makes one with LibreOffice if it’s installed. Otherwise ask the '
                      'user to export one into the source folder (PowerPoint: File › Export › PDF; Keynote: File › Export To › PDF), or run with '
                      '`--no-pdf` and tell them what that leaves out.')
        md += [f'3. `python3 tools/importdeck.py {slug}` converts every slide into `ct-slides.js` and `img/`, writes them into `{deck["entry"]}` with '
               'each slide’s title and speaker notes, and prints what came in.',
               f'4. Check it. `./sb shot {slug} <slide ids>` for a spread of slides (the first, the last, and any with charts, tables, or photos), '
               'and compare each with its original (`img/imp-NNN.jpg`, or the slide’s “original” button in the deck). Then `./sb audit '
               f'{slug}`. Fix what’s off in the slide’s shapes in `ct-slides.js`, and give a title to any slide the tool named “Slide N”.',
               f'5. `./sb export {slug}`, then `./sb commit {slug} "Imported from {Path(src).name}"`.',
               '6. Tell the user how many slides came in editable, which came in as pictures and why, and anything that looks different from the original.']
    else:
        md += [f'2. `python3 tools/importdeck.py {slug} --outline` writes `source/outline.md`: each slide’s title, words, tables, chart numbers, and notes, '
               'with a picture of each slide in `source/pages/`. ' + ('' if pdf else 'The pictures need a PDF of the deck in `source/` (or LibreOffice to make one); '
               'without one, ask the user to export it, or work from the words alone.'),
               f'3. Rebuild each slide in `{deck["entry"]}`, replacing the placeholder slide. Pick the layout that fits what the slide does (Add a slide’s basic '
               'layouts are in the deck already), keep its words, and put its speaker notes in `notes`. Keep the order, and make each section of the '
               'file a part (CATS). Rebuild a chart from its numbers in the outline as a FIGS figure. A slide that can’t be rebuilt well, like a dense '
               f'diagram or a photo collage, can stay as it was: `python3 tools/importdeck.py {slug} --slides 7,12 --no-wire` converts just those, '
               'and `ct(\'s7\', part, title, { notes })` puts one in LIBRARY.',
               '4. Keep the user’s titles. Where one is a label rather than a sentence, suggest an action title in that slide’s notes instead of changing it.',
               f'5. Check it: `./sb shot {slug} <slide ids>` against the pictures in `source/pages/`, and `./sb audit {slug}`. Everything passes Checks › Guides.',
               f'6. `./sb export {slug}`, then `./sb commit {slug} "Rebuilt from {Path(src).name}"`.',
               '7. Tell the user which layouts the slides went into, any chart or slide you kept as it was, anything you left out, and the titles you’d change.']
    return '\n'.join(md) + '\n'


def import_prompt(deck, src, mode, brief):
    where = f'It is set up at {deck["url"]} ({deck["folder"]}, slug {deck["slug"]}), with the file in source/.'
    steps = 'Follow the steps in BRIEF.md there (the Decks skill has the same under “Importing a deck”), and ask me before installing any Python packages.'
    if mode == 'keep':
        notes = '' if Path(src).suffix.lower() == '.pdf' else ', with its speaker notes'
        return (f'Import “{deck["name"]}” into Decks from {Path(src).name}, keeping each slide’s look: every slide converted so each word on it is '
                f'editable{notes}. {where} {steps} When it’s done, tell me how many slides came in editable, which came in as '
                'pictures and why, and anything that looks different from the original.')
    look = f' and the {brief["style"]} color set' if brief.get('style') else ''
    return (f'Import “{deck["name"]}” into Decks from {Path(src).name}, rebuilt in Decks’ own layouts{look}: the same slides in the same order, '
            f'with my words and my speaker notes. {where} {steps} Mark anything you rewrote in the notes. When it’s done, tell me which layouts '
            'the slides went into, any slide you kept as it was, and the titles you’d change.')


def made_root(deck):
    """The folder this app made for the deck (<decks folder>/<slug>), or None for a deck that lives in a
    folder of the user's own: only a made one may ever go to the Trash."""
    if not deck.get('template'):
        return None
    folder = Path(os.path.expanduser(deck['folder'])).resolve()
    if folder.parent != Path(sbconfig.DECKS).expanduser().resolve() or folder.name != deck['slug']:
        return None
    return folder


def to_trash(path):
    """Move a folder to the system Trash: ~/.Trash on macOS, the freedesktop Trash on Linux, else
    a .trash folder in the decks folder. Never deleted outright."""
    path = Path(path)
    stamp = time.strftime('%Y-%m-%d %H.%M.%S')
    if sys.platform == 'darwin':
        trash, info = Path.home() / '.Trash', None
    elif sys.platform.startswith('linux'):
        base = Path(os.environ.get('XDG_DATA_HOME') or Path.home() / '.local' / 'share') / 'Trash'
        trash, info = base / 'files', base / 'info'
    else:
        trash, info = Path(sbconfig.DECKS).expanduser() / '.trash', None
    trash.mkdir(parents=True, exist_ok=True)
    dest = trash / path.name
    if dest.exists():
        dest = trash / f'{path.name} {stamp}'
    shutil.move(str(path), str(dest))
    if info:   # so the file manager can put it back
        info.mkdir(parents=True, exist_ok=True)
        (info / (dest.name + '.trashinfo')).write_text(
            f'[Trash Info]\nPath={path}\nDeletionDate={time.strftime("%Y-%m-%dT%H:%M:%S")}\n', encoding='utf-8')
    return dest


def manage_deck(slug, op, registry_path=None):
    """hide / show: on or off the dashboard, nothing else changes.
    delete: leaves the registry; a deck this app made also moves its folder to the Trash
    (never deleted outright), while a deck in a folder of the user's own keeps every file where it is."""
    if op not in ('hide', 'show', 'delete'):
        raise ValueError(f'unknown op {op!r}')
    registry_path = Path(registry_path or sbconfig.REGISTRY)
    with _REG:
        reg = json.loads(registry_path.read_text(encoding='utf-8'))
        d = next((x for x in reg['decks'] if x.get('slug') == slug), None)
        if not d:
            raise KeyError(slug)
        trashed = None
        if op == 'hide':
            d['hidden'] = True
        elif op == 'show':
            d.pop('hidden', None)
        else:
            root = made_root(d)
            if root and root.exists():
                trashed = to_trash(root)
            reg['decks'] = [x for x in reg['decks'] if x is not d]
        tmp = registry_path.with_suffix('.json.tmp')
        tmp.write_text(json.dumps(reg, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
        os.replace(tmp, registry_path)
    return {'slug': slug, 'op': op, 'trashed': sbconfig.tilde(trashed) if trashed else None}


def create_deck(name, tid, brief=None, pieces=None, logo=None, registry_path=None, source=None):
    """source (Import a deck): {'path', 'name', 'mode': 'keep' or 'rebuild', 'move'}. The file goes into the deck's
    source/ folder (moved there with move, for a file the server just received), and the deck starts from the
    _import template. Keeping the look takes no color set or logo: the slides bring their own."""
    name = (name or '').strip()
    if not name:
        raise ValueError('Enter a name for the deck.')
    brief = {k: str(v).strip() for k, v in (brief or {}).items() if v and str(v).strip()}
    src_path = src_name = mode = None
    if source:
        src_path, src_name, mode = check_source(source)
        tid, pieces = IMPORT, []
        if mode == 'keep':
            brief = {k: v for k, v in brief.items() if k not in ('style', 'titles')}
            logo = None
    pieces = [{'id': re.sub(r'[^a-z0-9-]', '', str(pc.get('id', '')).lower()) or f'piece{i}', 'name': str(pc.get('name', '')).strip() or 'Version',
               'desc': str(pc.get('desc', '')).strip()} for i, pc in enumerate(pieces or []) if isinstance(pc, dict)]
    import base64
    if not source or mode == 'rebuild':   # house default: the LBMC style and logo unless the deck asks for another
        brief.setdefault('style', 'LBMC')
        if not logo and brief['style'] == 'LBMC':
            f = Path(__file__).parent / 'assets' / 'lbmc-logo.jpg'
            logo = {'data': 'data:image/jpeg;base64,' + base64.b64encode(f.read_bytes()).decode()}
    logo_raw = logo_ext = None
    if logo and logo.get('data'):   # the deck's own logo, from the form (a data URL): checked before anything is made
        m = re.match(r'data:image/(png|jpeg|jpg|svg\+xml|webp);base64,(.*)$', str(logo['data']), re.S)
        if not m:
            raise ValueError('The logo has to be a PNG, JPG, SVG, or WebP file.')
        logo_raw = base64.b64decode(m.group(2))
        if len(logo_raw) > 3_000_000:
            raise ValueError('That logo file is over 3 MB. Try a smaller one.')
        logo_ext = {'jpeg': 'jpg', 'svg+xml': 'svg'}.get(m.group(1), m.group(1))
    tmpl = template(tid)
    registry_path = Path(registry_path or sbconfig.REGISTRY)
    decks_dir = registry_path.parent
    with _REG:
        reg = json.loads(registry_path.read_text(encoding='utf-8')) if registry_path.exists() else {'decks': []}
        if any(d['name'].strip().lower() == name.lower() for d in reg['decks']):
            raise ValueError(f'There is already a deck called “{name}”. Try another name.')
        base = slugify(name); slug, k = base, 2
        while any(d['slug'] == slug for d in reg['decks']) or (decks_dir / slug).exists():
            slug, k = f'{base}-{k}', k + 1   # a different name that makes the same address: number it
        folder = decks_dir / slug
        prefix = new_prefix(name, [p for d in reg['decks'] for p in d.get('prefixes', [])] + list(_RESERVED))
        _RESERVED.add(prefix)
        folder.mkdir(parents=True)   # reserves the address
    moved = None
    try:   # if building fails, take back the reserved folder and prefix (and put a moved source file back)
        entry = tmpl.get('entry', 'deck.html')

        src = TEMPLATES / tid
        for p in src.iterdir():
            if p.name in ('template.json', 'content.js', 'pitches.js', 'README.md') or p.name.startswith('.'):
                continue
            (shutil.copytree if p.is_dir() else shutil.copy2)(p, folder / p.name)
        shared = TEMPLATES / '_engine' / 'img'
        if shared.exists():
            shutil.copytree(shared, folder / 'img', dirs_exist_ok=True)
        theme = theme_key(brief.get('style'))
        logo_path = None
        if logo_raw:
            (folder / 'img').mkdir(exist_ok=True)
            (folder / 'img' / f'logo.{logo_ext}').write_bytes(logo_raw)
            logo_path = f'img/logo.{logo_ext}'
        serif = 'serif' in brief.get('titles', '').lower() and 'sans' not in brief.get('titles', '').lower()
        if theme:
            recolor_images(folder, theme)
        (folder / entry).write_text(assemble(tid, name, prefix, pieces, theme, serif, logo_path, originals=bool(source)), encoding='utf-8')
        (folder / '.gitignore').write_text(GITIGNORE + ('source/\n' if source else ''), encoding='utf-8')
        if source:   # the file being imported, kept as it came
            (folder / 'source').mkdir()
            dest = folder / 'source' / src_name
            size = src_path.stat().st_size
            if source.get('move'):
                shutil.move(str(src_path), str(dest))
                moved = (dest, src_path)
            else:
                shutil.copy2(src_path, dest)

        deck = {'slug': slug, 'name': name, 'group': 'active', 'kind': tmpl.get('kind', 'workshop'),
                'folder': sbconfig.tilde(folder), 'entry': entry,
                'prefixes': [prefix], 'exportName': name, 'template': tid}
        if pieces:
            deck['pieces'] = [{'id': pc['id'], 'name': pc['name']} for pc in pieces]
        if source:
            deck.update({'source': f'source/{src_name}', 'importMode': mode, 'sourceSize': size_label(size)})
        deck['url'] = f'http://localhost:{sbconfig.PORT}/d/{slug}/{entry}'
        if source:
            (folder / 'BRIEF.md').write_text(import_brief(dict(deck, logo=logo_path), f'source/{src_name}', mode, brief), encoding='utf-8')
            deck['prompt'] = import_prompt(deck, src_name, mode, brief)
        else:
            md = [f'# {name}', '', f'Started {time.strftime("%B %-d, %Y")} from {tmpl_phrase(tmpl)}.', '']
            if pieces:
                md += ['## What to build', '', *[f'- **{pc["name"]}**' + (f': {pc["desc"]}' if pc.get('desc') else '') + (' (the main deck)' if i == 0 else '') for i, pc in enumerate(pieces)], '']
            for k, label in BRIEF_FIELDS:
                if brief.get(k):
                    md += [f'## {label}', '', brief[k], '']
            (folder / 'BRIEF.md').write_text('\n'.join(md), encoding='utf-8')
            deck['prompt'] = prompt_for(deck, brief, tmpl, pieces)
        with _REG:   # re-read: other decks may have been added while this one was built
            reg = json.loads(registry_path.read_text(encoding='utf-8')) if registry_path.exists() else {'decks': []}
            reg['decks'].insert(0, {k: v for k, v in deck.items() if k != 'url'})
            tmp = registry_path.with_suffix('.json.tmp')
            tmp.write_text(json.dumps(reg, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
            os.replace(tmp, registry_path)
            _RESERVED.discard(prefix)

    except BaseException:
        if moved and moved[0].exists():
            shutil.move(str(moved[0]), str(moved[1]))
        shutil.rmtree(folder, ignore_errors=True)
        _RESERVED.discard(prefix)
        raise
    try:   # version history for the deck; skipped without git
        git = lambda *a: subprocess.run(['git', *a], cwd=folder, capture_output=True, text=True)
        git('init', '-q')
        git('add', '-A')
        git('commit', '-q', '-m', f'New deck imported from {src_name}' if source else f'New deck from the {tmpl["name"]} template')
    except OSError:
        pass
    return deck



def create_demo(registry_path=None):
    """The deck a fresh install starts with: the talk template, with a brief about trying Decks."""
    return create_deck('Demo talk', 'talk', {
        'audience': 'people trying Decks for the first time',
        'point': 'Comment on a slide, ask your agent, and read its reply under the slide.',
        'length': '20 min',
        'style': 'Navy',
        'notes': 'A practice deck. Leave a comment under any slide, then click Ask Claude and paste the prompt into Claude Code or Codex.',
    }, registry_path=registry_path)


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else ''
    if cmd == 'templates':
        for t in list_templates():
            print(f'{t["id"]:14} {t["name"]}: {t.get("description", "")}')
    elif cmd == 'build':
        sys.stdout.write(assemble(sys.argv[2], 'Test deck', 'tst_'))
    else:
        print(__doc__)
