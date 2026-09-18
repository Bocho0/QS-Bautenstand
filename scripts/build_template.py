"""
Einmaliges Umbau-Skript: nimmt die ORIGINALE vorlage.docx und backt alle
statischen Formatierungs-/Strukturänderungen aus dem Muster-Handoff fest
in die Datei ein (Schriften, Farben, Ränder, Linien statt Rahmen,
Spaltenbreiten, verbundene Kopfzellen, Wetterband, Ampel-Legende, Fix des
Deckblatt-Kopfzeilenfehlers). Wird NUR einmal ausgeführt, um die neue
vorlage.docx zu erzeugen - der laufende Report-Code (report_core.py)
befüllt diese neue Vorlage danach nur noch mit Werten.
"""
import copy, sys
import docx
from docx.shared import Pt, Twips, RGBColor
from docx.oxml.ns import qn
from docx.oxml import parse_xml, OxmlElement
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

COL_TEXT = '1C1C1C'
COL_LABEL_COVER = '7A7A7A'
COL_LABEL_DOC = '8A8A8A'
COL_TABLE_HEAD = '6A6A6A'
COL_MUTED_ID = '5A5A5A'
COL_SUBHEAD = '4A4A4A'
COL_LINE_DARK = '000000'
COL_LINE_LIGHT = 'DEDEDE'
COL_BAND_BG = 'F3F3F1'

RT_COL_WIDTHS = [1021, 3657, 964, 964, 1247, 454, 198, 1361]
DOC_COL_WIDTHS = [510, 4590, 4766]
VT_COL_WIDTHS = [2154, 1124, 2821, 907, 2860]  # Name, Firma(gridSpan2), Kürzel, Email

LABEL_TAB_DXA = 851  # 15mm - Feststellungs-Raster
COVER_TAB_DXA = 1474  # 26mm - Deckblatt-Felder

AMPEL_LEGEND = [
    ('B2CB7F', 'Termin unkritisch'),
    ('F8A764', 'Termin kritisch'),
    ('F95649', 'Termin überschritten'),
]


def qd(tag):
    return f'{{{W}}}{tag}'


def set_run(run, size_pt=None, color=None, bold=None, name='Barlow', tracking=None, caps=None, underline=None):
    if name:
        run.font.name = name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.font.bold = bold
    if underline is not None:
        run.font.underline = underline
    rPr = run._r.get_or_add_rPr()
    if tracking is not None:
        old = rPr.find(qn('w:spacing'))
        if old is not None:
            rPr.remove(old)
        el = rPr.makeelement(qn('w:spacing'), {})
        el.set(qn('w:val'), str(tracking))
        rPr.append(el)
    if caps is not None:
        old = rPr.find(qn('w:caps'))
        if old is not None:
            rPr.remove(old)
        el = rPr.makeelement(qn('w:caps'), {})
        el.set(qn('w:val'), '1' if caps else '0')
        rPr.append(el)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = tcPr.makeelement(qn('w:tcBorders'), {})
        tcPr.append(borders)
    for side, spec in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if spec is None:
            continue
        old = borders.find(qn(f'w:{side}'))
        if old is not None:
            borders.remove(old)
        el = borders.makeelement(qn(f'w:{side}'), {})
        if spec == 'none':
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        else:
            val, sz, color = spec
            el.set(qn('w:val'), val); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        borders.append(el)


def clear_table_borders(table):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = tblPr.makeelement(qn('w:tblBorders'), {})
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = borders.makeelement(qn(f'w:{side}'), {})
        el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        borders.append(el)
    tblPr.append(borders)


def set_cell_margins(table, top=None, bottom=None, left=None, right=None):
    tblPr = table._tbl.tblPr
    mar = tblPr.find(qn('w:tblCellMar'))
    if mar is None:
        mar = tblPr.makeelement(qn('w:tblCellMar'), {})
        tblPr.append(mar)
    for side, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if val is None:
            continue
        old = mar.find(qn(f'w:{side}'))
        if old is not None:
            mar.remove(old)
        el = mar.makeelement(qn(f'w:{side}'), {})
        el.set(qn('w:w'), str(val)); el.set(qn('w:type'), 'dxa')
        mar.append(el)


def set_cell_margin(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = tcPr.find(qn('w:tcMar'))
    if mar is None:
        mar = tcPr.makeelement(qn('w:tcMar'), {})
        tcPr.append(mar)
    for side, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if val is None:
            continue
        old = mar.find(qn(f'w:{side}'))
        if old is not None:
            mar.remove(old)
        el = mar.makeelement(qn(f'w:{side}'), {})
        el.set(qn('w:w'), str(val)); el.set(qn('w:type'), 'dxa')
        mar.append(el)


def set_verteiler_widths(vt, name_w, firma_w, kuerzel_w, email_w):
    """Die Verteiler-Tabelle hat 5 Rasterspalten, wobei 'Firma' in JEDER
    Zeile (Kopf wie Daten) bereits als EINE Zelle mit gridSpan=2 über die
    Rasterspalten 1+2 verschmolzen ist. Die generische set_col_width_dxa()
    wuerde bei separatem Aufruf fuer Index 1 und 2 zweimal dieselbe Zelle
    treffen (zweiter Aufruf ueberschreibt den ersten) - deshalb hier eine
    eigene, auf diese Struktur zugeschnittene Breitenzuweisung."""
    grid = vt._tbl.tblGrid
    cols = grid.findall(qn('w:gridCol'))
    firma_ratio = 0.2848  # Aufteilungsverhaeltnis aus der Original-Vorlage
    firma_p1 = int(round(firma_w * firma_ratio))
    firma_p2 = firma_w - firma_p1
    for i, wdt in enumerate([name_w, firma_p1, firma_p2, kuerzel_w, email_w]):
        if i < len(cols):
            cols[i].set(qn('w:w'), str(wdt))
    for row in vt.rows:
        tcs = row._tr.findall(qn('w:tc'))
        if len(tcs) != 4:
            continue
        for tc, wdt in zip(tcs, [name_w, firma_w, kuerzel_w, email_w]):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = tc.makeelement(qn('w:tcPr'), {})
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = tcPr.makeelement(qn('w:tcW'), {})
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(wdt)); tcW.set(qn('w:type'), 'dxa')


def set_col_width_dxa(table, col_idx, dxa):
    for row in table.rows:
        if col_idx < len(row.cells):
            tc = row.cells[col_idx]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = tcPr.makeelement(qn('w:tcW'), {})
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(dxa)); tcW.set(qn('w:type'), 'dxa')
    grid = table._tbl.tblGrid
    cols = grid.findall(qn('w:gridCol'))
    if col_idx < len(cols):
        cols[col_idx].set(qn('w:w'), str(dxa))


def make_tab_stop_xml(pos_dxa):
    return f'<w:tabs xmlns:w="{W}"><w:tab w:val="left" w:pos="{pos_dxa}"/></w:tabs>'


def set_para_tab_and_indent(paragraph, tab_dxa, hanging=True):
    pPr = paragraph._p.get_or_add_pPr()
    old_tabs = pPr.find(qn('w:tabs'))
    if old_tabs is not None:
        pPr.remove(old_tabs)
    tabs_el = parse_xml(make_tab_stop_xml(tab_dxa))
    pPr.append(tabs_el)
    old_ind = pPr.find(qn('w:ind'))
    if old_ind is not None:
        pPr.remove(old_ind)
    ind = pPr.makeelement(qn('w:ind'), {})
    ind.set(qn('w:left'), str(tab_dxa))
    if hanging:
        ind.set(qn('w:hanging'), str(tab_dxa))
    pPr.append(ind)


def main(src_path, out_path):
    doc = docx.Document(src_path)

    # ------------------------------------------------------------------
    # 1) Seitenränder
    # ------------------------------------------------------------------
    for sec in doc.sections:
        sec.top_margin = Twips(1134)
        sec.bottom_margin = Twips(737)
        sec.left_margin = Twips(1021)
        sec.right_margin = Twips(1021)
    for tbl in doc.tables:
        tblPr = tbl._tbl.tblPr
        ind = tblPr.find(qn('w:tblInd'))
        if ind is not None:
            ind.set(qn('w:w'), '0')

    verteiler_table = doc.tables[0]
    wetter_table = doc.tables[1]
    rahmentermine_table = doc.tables[2]
    dokumentation_table = doc.tables[3]

    # ------------------------------------------------------------------
    # 2) Kopfzeile: Deckblatt-Fehler beheben (fehlende Unterzeile + Extra-
    #    Leerzeile) + Typografie
    # ------------------------------------------------------------------
    sec = doc.sections[0]
    hdr, fph = sec.header, sec.first_page_header

    fph_paras = list(fph.paragraphs)
    if fph_paras:
        first = fph_paras[0]
        has_drawing = first._p.find('.//' + qn('w:drawing')) is not None
        if not has_drawing and not first.text.strip():
            first._p.getparent().remove(first._p)

    has_subtitle = any(('Leistungsfeststellung' in p.text or 'Qualitätssicherung' in p.text)
                        for p in fph.paragraphs)
    if not has_subtitle:
        subtitle_src = None
        for p in hdr.paragraphs:
            if 'Leistungsfeststellung' in p.text or 'Qualitätssicherung' in p.text:
                subtitle_src = p._p
                break
        project_p = None
        for p in fph.paragraphs:
            if p._p.find('.//' + qn('w:drawing')) is not None or 'XY' in p.text:
                project_p = p._p
        if subtitle_src is not None and project_p is not None:
            clone = copy.deepcopy(subtitle_src)
            project_p.addnext(clone)

    for section_hdr in (hdr, fph):
        for p in section_hdr.paragraphs:
            pPr = p._p.find(qn('w:pPr'))
            if pPr is not None:
                pbdr = pPr.find(qn('w:pBdr'))
                if pbdr is not None:
                    top = pbdr.find(qn('w:top'))
                    if top is not None:
                        top.set(qn('w:sz'), '6')
                        top.set(qn('w:color'), COL_LINE_DARK)
            txt = p.text
            if 'Leistungsfeststellung' in txt or 'Qualitätssicherung' in txt:
                for r in p.runs:
                    if r.text.strip():
                        set_run(r, size_pt=9, color=COL_SUBHEAD)
            elif txt.strip():
                for r in p.runs:
                    if r.text.strip() and r._r.find(qn('w:drawing')) is None:
                        set_run(r, tracking=6)
            # Die Projektzeile enthaelt in der Original-Vorlage mehrere
            # (leere) Runs mit demselben Logo als Drawing - dadurch wird
            # das Logo mehrfach uebereinander/nebeneinander gerendert.
            # Nur den ERSTEN Drawing-Run behalten.
            drawing_runs = [r for r in p.runs if r._r.find(qn('w:drawing')) is not None]
            for extra in drawing_runs[1:]:
                extra._r.getparent().remove(extra._r)

    # ------------------------------------------------------------------
    # 2b) Halbgeviertstrich statt Bindestrich in den statischen CI-Texten
    #     (Muster nutzt durchgängig "–", nicht "-").
    # ------------------------------------------------------------------
    def fix_dash_runs(paragraph):
        runs = list(paragraph.runs)
        for i, r in enumerate(runs):
            if ' - ' in r.text:
                r.text = r.text.replace(' - ', ' – ')
            elif r.text.strip() == '-':
                r.text = r.text.replace('-', '–', 1)
            elif r.text.startswith('- '):
                r.text = '– ' + r.text[2:]
            elif r.text.endswith('-') and i + 1 < len(runs) and runs[i + 1].text.startswith(' '):
                r.text = r.text[:-1] + '–'
            elif r.text.endswith('- ') and (i == 0 or runs[i - 1].text.endswith(' ') or runs[i - 1].text == ''):
                pass  # bereits durch Nachbarlauf behandelt

    for p in doc.paragraphs:
        ptext = p.text
        if ('Leistungsfeststellung' in ptext or 'LEISTUNGSFES' in ptext.upper()
                or 'Qualitätssicherung' in ptext or 'QUALITÄTSSICHERUNG' in ptext.upper()):
            fix_dash_runs(p)
    for sec_hdr in (hdr, fph):
        for p in sec_hdr.paragraphs:
            if 'Leistungsfeststellung' in p.text or 'Qualitätssicherung' in p.text:
                fix_dash_runs(p)

    # ------------------------------------------------------------------
    # 3) Deckblatt-Titel
    # ------------------------------------------------------------------
    for p in doc.paragraphs:
        if p.text.strip().upper().startswith('LEISTUNGSFES'):
            for r in p.runs:
                set_run(r, size_pt=12, color=COL_TEXT, bold=True, tracking=62)
            break

    # ------------------------------------------------------------------
    # 4) Deckblatt-Felder: Unterstreichung/Doppelpunkt entfernen, Label
    #    grau, Wert per Tabulator bei 26mm ausrichten
    # ------------------------------------------------------------------
    cover_labels = ['Thema', 'Verfasser', 'Datum', 'Seitenanzahl', 'Kunde', 'Verteilung', 'Anlagen']
    for p in doc.paragraphs:
        stripped = p.text.strip()
        matched = None
        for lbl in cover_labels:
            if stripped.startswith(lbl + ':'):
                matched = lbl
                break
        if not matched or not p.runs:
            continue
        value = stripped[len(matched) + 1:].strip()
        base_run = p.runs[0]
        for extra in list(p.runs)[1:]:
            extra._r.getparent().remove(extra._r)
        r1 = p.runs[0]
        r1.text = matched
        set_run(r1, size_pt=9, color=COL_LABEL_COVER, underline=False, bold=False)
        r2 = p.add_run('\t' + value)
        set_run(r2, size_pt=9, color=COL_TEXT)
        p.paragraph_format.left_indent = Twips(0)
        p.paragraph_format.first_line_indent = Twips(0)
        set_para_tab_and_indent(p, COVER_TAB_DXA, hanging=False)
        # links kein Haengeeinzug fuer Cover-Felder (nur Tab), daher
        # ind.left auf 0 zuruecksetzen (set_para_tab_and_indent setzt sonst
        # links=tab_dxa)
        pPr = p._p.find(qn('w:pPr'))
        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            ind.set(qn('w:left'), '0')
            if ind.get(qn('w:hanging')) is not None:
                del ind.attrib[qn('w:hanging')]
        # Einzelne Absaetze (z.B. "Kunde:") nutzen in der Original-Vorlage
        # einen abweichenden Absatzstil ("altTitel") mit eigenem, teils
        # negativem Einzug, der die direkte Formatierung oben nur teilweise
        # überstimmt - Stil-Override entfernen, damit alle Cover-Felder
        # garantiert identisch (Standardstil) ausgerichtet sind.
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            pPr.remove(pStyle)
        jc = pPr.find(qn('w:jc'))
        if jc is not None:
            pPr.remove(jc)

    for p in doc.paragraphs:
        if p.text.strip() in ('Mustermannstr. 1', '12345 Hausen'):
            for r in p.runs:
                r.text = ''

    # ------------------------------------------------------------------
    # 4b) Vertikale Abstände auf dem Deckblatt exakt nach den mm-Werten aus
    #     dem Muster-HTML setzen (bislang unkontrolliert = Word-Standard).
    #     mm -> dxa: *56.6929
    # ------------------------------------------------------------------
    def set_space_before(paragraph, dxa):
        pPr = paragraph._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = pPr.makeelement(qn('w:spacing'), {})
            pPr.append(sp)
        sp.set(qn('w:before'), str(dxa))

    def set_space_after(paragraph, dxa):
        pPr = paragraph._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = pPr.makeelement(qn('w:spacing'), {})
            pPr.append(sp)
        sp.set(qn('w:after'), str(dxa))

    SPACING_BEFORE = {
        'Thema': 680,        # 12mm - Abstand nach dem Titel
        'Verfasser': 170,    # 3mm  - Abstand nach "Thema"
        'Kunde': 170,        # 3mm  - Abstand nach "Seitenanzahl"
        'Verteilung': 340,   # 6mm  - Abstand nach der Verteiler-Tabelle
    }
    for p in doc.paragraphs:
        stripped = p.text.strip()
        for lbl, dxa in SPACING_BEFORE.items():
            if stripped == lbl or (p.runs and p.runs[0].text.strip() == lbl):
                set_space_before(p, dxa)
                break
        if stripped.upper().startswith('LEISTUNGSFES'):
            set_space_before(p, 1247)  # 22mm - Abstand nach der Kopfzeilenlinie
        if stripped.rstrip(':') == 'Verteiler':
            set_space_before(p, 623)   # 11mm - Abstand nach "Kunde"
            set_space_after(p, 113)    # 2mm  - Abstand vor der Verteiler-Tabelle
        if stripped in ('Rahmentermine',):
            set_space_before(p, 454)   # 8mm - Abstand nach dem Wetterband
        if stripped in ('Dokumentation',):
            set_space_before(p, 397)   # 7mm - Abstand nach der Kopfzeilenlinie
            set_space_after(p, 198)    # 3,5mm - Abstand vor der Dokumentationstabelle
        if stripped.startswith('Hinweis:'):
            set_space_before(p, 45)    # 0,8mm - dicht unter "Rahmentermine"
            set_space_after(p, 227)    # 4mm - Abstand vor der Rahmentermine-Tabelle

    # Leerzeilen zwischen einer Überschrift/einem Hinweis und der jeweils
    # folgenden Tabelle sollen keinen zusätzlichen eigenen Abstand mehr
    # beitragen (der wird jetzt über space_after der Überschrift gesteuert)
    # - sonst addiert sich beides.
    def collapse_blank_before_table(label_predicate):
        for p in doc.paragraphs:
            if label_predicate(p.text.strip()):
                nxt = p._p.getnext()
                if nxt is not None and nxt.tag == qn('w:p'):
                    from docx.text.paragraph import Paragraph
                    nxt_p = Paragraph(nxt, p._parent)
                    if not nxt_p.text.strip():
                        set_space_before(nxt_p, 0)
                        set_space_after(nxt_p, 0)
                        pPr = nxt_p._p.get_or_add_pPr()
                        old_rpr = pPr.find(qn('w:rPr'))
                        sz_el = old_rpr.find(qn('w:sz')) if old_rpr is not None else None
                        if sz_el is not None:
                            sz_el.set(qn('w:val'), '2')
                break

    collapse_blank_before_table(lambda t: t.startswith('Hinweis:'))
    collapse_blank_before_table(lambda t: t == 'Dokumentation')

    # ------------------------------------------------------------------
    # 5) "Verteiler" Eyebrow + Tabelle
    # ------------------------------------------------------------------
    for p in doc.paragraphs:
        if p.text.strip().rstrip(':') == 'Verteiler':
            if p.runs:
                base_run = p.runs[0]
                for extra in list(p.runs)[1:]:
                    extra._r.getparent().remove(extra._r)
                r = p.runs[0]
                r.text = 'Verteiler'
                set_run(r, size_pt=7.5, color=COL_TABLE_HEAD, tracking=24, caps=True,
                        underline=False, bold=False)
            break

    vt = verteiler_table
    header_base = None
    for c in vt.rows[0].cells:
        for p in c.paragraphs:
            if p.runs:
                header_base = p.runs[0]; break
        if header_base is not None:
            break
    clear_table_borders(vt)
    for c in vt.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs:
                set_run(r, size_pt=7.5, color=COL_TABLE_HEAD, tracking=4)
        set_cell_borders(c, top='none', left='none', right='none',
                          bottom=('single', 6, COL_LINE_DARK))
    set_verteiler_widths(vt, 2154, 3945, 907, 2860)
    # Zeile 1 ist strukturell defekt (nur 1 Zelle mit gridSpan statt der
    # ueblichen 4 Zellen, siehe Kommentar oben bei set_verteiler_widths) -
    # durch eine leere Kopie der intakten Zeile 2 ersetzen.
    if len(vt.rows) > 2:
        row1_tr = vt.rows[1]._tr
        row2_clone = copy.deepcopy(vt.rows[2]._tr)
        for tc in row2_clone.findall(qn('w:tc')):
            for p_el in tc.findall(qn('w:p')):
                for r_el in p_el.findall(qn('w:r')):
                    for t_el in r_el.findall(qn('w:t')):
                        t_el.text = ''
        row1_tr.addprevious(row2_clone)
        row1_tr.getparent().remove(row1_tr)
    # Datenzeilen: duenne graue Linie statt Punktrahmen (Muster Abschnitt 4)
    for row in list(vt.rows)[1:]:
        # Zeile 1 der Vorlage hatte eine versteckte trPr-Eigenschaft
        # (gridAfter/wAfter), die Platz fuer "unsichtbare" Spalten reserviert
        # und dadurch die Linie/Zellen dieser einen Zeile gegenueber allen
        # anderen verkuerzt hat - dauerhaft entfernen.
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is not None:
            for tag in ('w:gridAfter', 'w:wAfter', 'w:gridBefore', 'w:wBefore'):
                el = trPr.find(qn(tag))
                if el is not None:
                    trPr.remove(el)
        for cell in row.cells:
            set_cell_borders(cell, top='none', left='none', right='none',
                              bottom=('single', 3, COL_LINE_LIGHT))
            for p in cell.paragraphs:
                for r in p.runs:
                    set_run(r, size_pt=9, color=COL_TEXT)

    # ------------------------------------------------------------------
    # 6) Wetter-Tabelle entfernen, durch Wetterband-Absatz ersetzen
    #    (3 Runs per Tab getrennt - Platzhaltertexte, werden zur Laufzeit
    #    per Textersetzung befuellt)
    # ------------------------------------------------------------------
    band_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr>
        <w:tabs><w:tab w:val="left" w:pos="1600"/><w:tab w:val="left" w:pos="3400"/></w:tabs>
        <w:shd w:val="clear" w:color="auto" w:fill="{COL_BAND_BG}"/>
        <w:spacing w:before="397" w:after="0"/>
        <w:ind w:left="90" w:right="90"/>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="3A3A3A"/><w:b/></w:rPr>
      <w:t>WETTERBAND_DATUM</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="3A3A3A"/></w:rPr><w:tab/>
      <w:t>WETTERBAND_MIN</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:color w:val="3A3A3A"/></w:rPr><w:tab/>
      <w:t>WETTERBAND_MAX</w:t></w:r>
    </w:p>'''
    band_p = parse_xml(band_xml)
    tbl_el = wetter_table._tbl
    tbl_el.addprevious(band_p)
    tbl_el.getparent().remove(tbl_el)

    # ------------------------------------------------------------------
    # 7) Rahmentermine-Tabelle: Kopf verbinden, Formatierung, auf EINE
    #    Datenzeile reduzieren (wird zur Laufzeit pro Eintrag geklont -
    #    wie die Dokumentationstabelle), Ampel-Legende einfuegen
    # ------------------------------------------------------------------
    rt = rahmentermine_table
    header_tcs = rt.rows[0]._tr.findall(qn('w:tc'))
    if len(header_tcs) >= 7:
        bautenstand_tc = header_tcs[3]
        tcPr = bautenstand_tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = bautenstand_tc.makeelement(qn('w:tcPr'), {})
            bautenstand_tc.insert(0, tcPr)
        old_span = tcPr.find(qn('w:gridSpan'))
        if old_span is not None:
            tcPr.remove(old_span)
        span_el = tcPr.makeelement(qn('w:gridSpan'), {})
        span_el.set(qn('w:val'), '3')
        tcPr.insert(0, span_el)
        for redundant in header_tcs[4:6]:
            redundant.getparent().remove(redundant)

    rt_base_run = None
    for c in rt.rows[0].cells:
        for p in c.paragraphs:
            if p.runs:
                rt_base_run = p.runs[0]; break
        if rt_base_run is not None:
            break
    clear_table_borders(rt)
    for c in rt.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs:
                set_run(r, size_pt=7.5, color=COL_TABLE_HEAD, tracking=4)
                if r.text.rstrip().endswith(':'):
                    r.text = r.text.rstrip()[:-1]
        set_cell_borders(c, top='none', left='none', right='none',
                          bottom=('single', 6, COL_LINE_DARK))

    # Auf eine einzige Datenzeile reduzieren (Klonvorlage fuer
    # report_core.py); Beispielwerte durch generische Platzhalter ersetzen,
    # Haus-Spalte grau (gedaempft), rechtsbuendige %-Spalte.
    data_rows = list(rt.rows)[1:]
    if data_rows:
        keep_tr = data_rows[0]._tr
        for row in data_rows[1:]:
            rt._tbl.remove(row._tr)
        keep_row = rt.rows[1]
        for idx, cell in enumerate(keep_row.cells):
            for p in cell.paragraphs:
                for r in list(p.runs):
                    r.text = ''
                if p.runs:
                    r = p.runs[0]
                else:
                    r = p.add_run()
                r.text = '' if idx not in (0,) else ''
                color = COL_MUTED_ID if idx == 0 else COL_TEXT
                set_run(r, size_pt=9, color=color)
                if idx == 5:
                    from docx.enum.text import WD_ALIGN_PARAGRAPH
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_cell_borders(cell, top='none', left='none', right='none',
                              bottom=('single', 3, COL_LINE_LIGHT))
        # Ampelfarbe der Klonvorlage leeren (weiss)
        if len(keep_row.cells) > 6:
            color_cell = keep_row.cells[6]
            tcPr = color_cell._tc.get_or_add_tcPr()
            old_shd = tcPr.find(qn('w:shd'))
            if old_shd is not None:
                tcPr.remove(old_shd)
            shd = tcPr.makeelement(qn('w:shd'), {})
            shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'FFFFFF')
            tcPr.append(shd)

    set_cell_margins(rt, top=48, bottom=48, left=57, right=57)
    for idx, width in enumerate(RT_COL_WIDTHS):
        set_col_width_dxa(rt, idx, width)

    # Ampel-Legende (rein statisch) direkt unter der Rahmentermine-Tabelle
    # einfuegen.
    swatch_w, gap_w, swatch_label_gap = 182, 397, 91
    label_widths = [1250, 1150, 1550]
    grid_widths, cells_xml = [], ''
    for k, ((color, label), lw) in enumerate(zip(AMPEL_LEGEND, label_widths)):
        grid_widths += [swatch_w, swatch_label_gap, lw]
        if k < len(AMPEL_LEGEND) - 1:
            grid_widths += [gap_w]
        cells_xml += (
            f'<w:tc><w:tcPr><w:tcW w:w="{swatch_w}" w:type="dxa"/>'
            f'<w:shd w:val="clear" w:color="auto" w:fill="{color}"/>'
            f'<w:tcBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'
            f'<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>'
        )
        cells_xml += (
            f'<w:tc><w:tcPr><w:tcW w:w="{swatch_label_gap}" w:type="dxa"/>'
            f'<w:tcBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'
            f'</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>'
        )
        cells_xml += (
            f'<w:tc><w:tcPr><w:tcW w:w="{lw}" w:type="dxa"/>'
            f'<w:tcBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'
            f'<w:vAlign w:val="center"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/>'
            f'<w:sz w:val="15"/><w:szCs w:val="15"/><w:color w:val="{COL_MUTED_ID}"/></w:rPr>'
            f'<w:t xml:space="preserve">{label}</w:t></w:r></w:p></w:tc>'
        )
        if k < len(AMPEL_LEGEND) - 1:
            cells_xml += (
                f'<w:tc><w:tcPr><w:tcW w:w="{gap_w}" w:type="dxa"/>'
                f'<w:tcBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'
                f'</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>'
            )
    grid_xml = ''.join(f'<w:gridCol w:w="{gw}"/>' for gw in grid_widths)
    legend_tbl_xml = f'''<w:tbl xmlns:w="{W}">
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblLayout w:type="fixed"/>
        <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
        <w:tblBorders>
          <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
          <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
          <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>{grid_xml}</w:tblGrid>
      <w:tr>{cells_xml}</w:tr>
    </w:tbl>'''
    legend_tbl = parse_xml(legend_tbl_xml)
    spacer_p = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:spacing w:before="198" w:after="0"/></w:pPr></w:p>')
    rt._tbl.addnext(legend_tbl)
    rt._tbl.addnext(spacer_p)

    # ------------------------------------------------------------------
    # 8) "Rahmentermine:"/"Dokumentation:" Abschnittstitel + "Hinweis:"
    # ------------------------------------------------------------------
    for p in doc.paragraphs:
        t = p.text.strip()
        if t in ('Rahmentermine:', 'Dokumentation:') and p.runs:
            base_run = p.runs[0]
            for extra in list(p.runs)[1:]:
                extra._r.getparent().remove(extra._r)
            r = p.runs[0]
            r.text = t.rstrip(':')
            set_run(r, size_pt=10.5, color=COL_TEXT, bold=True)
        elif t.startswith('Hinweis:') and p.runs:
            for r in p.runs:
                set_run(r, size_pt=8, color=COL_MUTED_ID, bold=False)

    # ------------------------------------------------------------------
    # 9) Dokumentationstabelle: Kopf, Spaltenbreiten, EINE Muster-
    #    Feststellungszeile im Label/Wert-Raster (Klonvorlage fuer
    #    report_core.py)
    # ------------------------------------------------------------------
    dt = dokumentation_table
    dt_base_run = None
    for c in dt.rows[0].cells:
        for p in c.paragraphs:
            if p.runs:
                dt_base_run = p.runs[0]; break
        if dt_base_run is not None:
            break
    clear_table_borders(dt)
    for c in dt.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs:
                set_run(r, size_pt=7.5, color=COL_TABLE_HEAD, tracking=4)
                if r.text.rstrip().endswith(':'):
                    r.text = r.text.rstrip()[:-1]
        set_cell_borders(c, top='none', left='none', right='none',
                          bottom=('single', 6, COL_LINE_DARK))
    set_cell_margins(dt, top=48, bottom=48, left=100, right=100)
    for c in dt.rows[0].cells:
        set_cell_margin(c, top=48, bottom=48, left=100, right=100)
    for idx, width in enumerate(DOC_COL_WIDTHS):
        set_col_width_dxa(dt, idx, width)

    # Musterzeile (Zeile 1) komplett neu aufbauen: Nr.-Zelle (grau, Platz-
    # halter leer), Feststellungs-Zelle mit EINER Prototyp-Label/Wert-Zeile
    # ("PROTO_LABEL" \t "PROTO_VALUE") + einer "Stand:"-Zeile, Foto-Zelle
    # unveraendert (Bild wird zur Laufzeit eingefuegt).
    sample_row = dt.rows[1]
    nr_cell, feststellung_cell, foto_cell = sample_row.cells[0], sample_row.cells[1], sample_row.cells[2]

    # Nr.-Zelle
    for p in list(nr_cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = nr_cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    p0.text = ''
    p0.paragraph_format.space_before = Pt(16)
    r_nr = p0.add_run('1')
    set_run(r_nr, size_pt=9, color=COL_MUTED_ID)

    # Feststellungs-Zelle: alle vorhandenen Absaetze entfernen, durch genau
    # 2 Prototyp-Absaetze ersetzen (Label/Wert-Zeile + Stand-Zeile), die
    # report_core.py als Kopiervorlage (deepcopy) verwendet.
    tc = feststellung_cell._tc
    for p_el in list(tc.findall(qn('w:p'))):
        tc.remove(p_el)
    proto_line_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr>
        <w:spacing w:before="320"/>
        <w:tabs><w:tab w:val="left" w:pos="{LABEL_TAB_DXA}"/></w:tabs>
        <w:ind w:left="{LABEL_TAB_DXA}" w:hanging="{LABEL_TAB_DXA}"/>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="{COL_LABEL_DOC}"/></w:rPr><w:t>PROTO_LABEL</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="{COL_TEXT}"/></w:rPr><w:tab/><w:t>PROTO_VALUE</w:t></w:r>
    </w:p>'''
    proto_indent_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr><w:ind w:left="{LABEL_TAB_DXA}"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="{COL_TEXT}"/></w:rPr><w:t>PROTO_CONT</w:t></w:r>
    </w:p>'''
    proto_stand_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr><w:spacing w:before="60"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="{COL_LABEL_DOC}"/></w:rPr><w:t>PROTO_STAND</w:t></w:r>
    </w:p>'''
    tc.append(parse_xml(proto_line_xml))
    tc.append(parse_xml(proto_indent_xml))
    tc.append(parse_xml(proto_stand_xml))

    # Zeile 1 ist die neue Klonvorlage fuer report_core.py - alle weiteren
    # (alten Beispiel-)Zeilen der Vorlage komplett entfernen.
    for row in list(dt.rows)[2:]:
        dt._tbl.remove(row._tr)

    # Foto-Zelle: bestehendes Beispielbild entfernen, leere Zelle mit
    # zentrierter Ausrichtung uebrig lassen (Bild kommt zur Laufzeit).
    tc_foto = foto_cell._tc
    old_valign = None
    tcPr = tc_foto.find(qn('w:tcPr'))
    if tcPr is not None:
        old_valign = tcPr.find(qn('w:vAlign'))
    for p_el in list(tc_foto.findall(qn('w:p'))):
        tc_foto.remove(p_el)
    tc_foto.append(parse_xml(f'<w:p xmlns:w="{W}"/>'))

    doc.save(out_path)
    print('Template gespeichert:', out_path)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
