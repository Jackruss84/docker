# -*- coding: utf-8 -*-
"""
Valoria — Générateur PDF Avis de Valeur Propriétaire
Thème : blanc, classique, typographie élégante
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, HRFlowable, KeepTogether,
    NextPageTemplate
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os

# ── Palette ──────────────────────────────────────────────────────────────────
BLACK      = colors.HexColor('#1a1a1a')
DARK_GRAY  = colors.HexColor('#2d2d2d')
MID_GRAY   = colors.HexColor('#666666')
LIGHT_GRAY = colors.HexColor('#999999')
RULE_GRAY  = colors.HexColor('#d4d4d4')
BG_LIGHT   = colors.HexColor('#f7f7f5')
BG_CARD    = colors.HexColor('#f2f2ef')
ACCENT     = colors.HexColor('#1a3a5c')   # bleu marine profond
ACCENT2    = colors.HexColor('#c8a96e')   # or discret
WHITE      = colors.white
GREEN_NUM  = colors.HexColor('#2d6a4f')
RED_NUM    = colors.HexColor('#9b2226')

W, H = A4
ML = 22*mm; MR = 22*mm; MT = 20*mm; MB = 20*mm
CW = W - ML - MR

MONTHS_FR = ['janvier','février','mars','avril','mai','juin',
             'juillet','août','septembre','octobre','novembre','décembre']

def fmt_price(p):
    try:
        n = int(float(p))
        s = f"{n:,}".replace(',', ' ')
        return f"{s} €"
    except:
        return str(p)

def date_fr():
    d = datetime.now()
    return f"{d.day} {MONTHS_FR[d.month-1]} {d.year}"

# ── Styles ────────────────────────────────────────────────────────────────────
def make_styles():
    S = {}

    # Cover
    S['agency_name'] = ParagraphStyle('agency_name',
        fontName='Helvetica-Bold', fontSize=10, textColor=ACCENT,
        alignment=TA_CENTER, leading=14, tracking=3)
    S['agency_sub'] = ParagraphStyle('agency_sub',
        fontName='Helvetica', fontSize=8.5, textColor=MID_GRAY,
        alignment=TA_CENTER, leading=12)
    S['cover_label'] = ParagraphStyle('cover_label',
        fontName='Helvetica', fontSize=9, textColor=ACCENT2,
        alignment=TA_CENTER, leading=12, tracking=5)
    S['cover_title'] = ParagraphStyle('cover_title',
        fontName='Helvetica-Bold', fontSize=32, textColor=ACCENT,
        alignment=TA_CENTER, leading=38)
    S['cover_address'] = ParagraphStyle('cover_address',
        fontName='Helvetica', fontSize=15, textColor=BLACK,
        alignment=TA_CENTER, leading=22)
    S['cover_address2'] = ParagraphStyle('cover_address2',
        fontName='Helvetica', fontSize=12, textColor=MID_GRAY,
        alignment=TA_CENTER, leading=17)
    S['cover_owner'] = ParagraphStyle('cover_owner',
        fontName='Helvetica-Bold', fontSize=12, textColor=DARK_GRAY,
        alignment=TA_CENTER, leading=17)
    S['cover_price'] = ParagraphStyle('cover_price',
        fontName='Helvetica-Bold', fontSize=36, textColor=ACCENT,
        alignment=TA_CENTER, leading=42)
    S['cover_price_label'] = ParagraphStyle('cover_price_label',
        fontName='Helvetica', fontSize=8, textColor=LIGHT_GRAY,
        alignment=TA_CENTER, leading=11, tracking=3)
    S['cover_meta'] = ParagraphStyle('cover_meta',
        fontName='Helvetica', fontSize=9, textColor=MID_GRAY,
        alignment=TA_CENTER, leading=14)

    # Body
    S['section_num'] = ParagraphStyle('section_num',
        fontName='Helvetica', fontSize=8, textColor=ACCENT2,
        alignment=TA_LEFT, leading=10, tracking=2)
    S['section_title'] = ParagraphStyle('section_title',
        fontName='Helvetica-Bold', fontSize=16, textColor=ACCENT,
        alignment=TA_LEFT, leading=20, spaceBefore=4, spaceAfter=6)
    S['h2'] = ParagraphStyle('h2',
        fontName='Helvetica-Bold', fontSize=12, textColor=DARK_GRAY,
        leading=16, spaceBefore=10, spaceAfter=4)
    S['h3'] = ParagraphStyle('h3',
        fontName='Helvetica-Bold', fontSize=10, textColor=MID_GRAY,
        leading=13, spaceBefore=6, spaceAfter=2, tracking=1)
    S['body'] = ParagraphStyle('body',
        fontName='Helvetica', fontSize=10, textColor=DARK_GRAY,
        leading=16, spaceAfter=8, alignment=TA_JUSTIFY)
    S['body_sm'] = ParagraphStyle('body_sm',
        fontName='Helvetica', fontSize=9, textColor=MID_GRAY,
        leading=13, spaceAfter=5)
    S['body_center'] = ParagraphStyle('body_center',
        fontName='Helvetica', fontSize=10, textColor=DARK_GRAY,
        leading=16, alignment=TA_CENTER)
    S['price_big'] = ParagraphStyle('price_big',
        fontName='Helvetica-Bold', fontSize=28, textColor=ACCENT,
        alignment=TA_CENTER, leading=34)
    S['price_label'] = ParagraphStyle('price_label',
        fontName='Helvetica', fontSize=8, textColor=LIGHT_GRAY,
        alignment=TA_CENTER, leading=10, tracking=2)
    S['price_range'] = ParagraphStyle('price_range',
        fontName='Helvetica-Bold', fontSize=13, textColor=ACCENT,
        alignment=TA_CENTER, leading=18)
    S['letter'] = ParagraphStyle('letter',
        fontName='Helvetica', fontSize=10.5, textColor=DARK_GRAY,
        leading=18, spaceAfter=10, alignment=TA_JUSTIFY)
    S['letter_bold'] = ParagraphStyle('letter_bold',
        fontName='Helvetica-Bold', fontSize=10.5, textColor=BLACK,
        leading=16, spaceAfter=4)
    S['sign'] = ParagraphStyle('sign',
        fontName='Helvetica', fontSize=10, textColor=MID_GRAY,
        leading=15)
    S['sign_bold'] = ParagraphStyle('sign_bold',
        fontName='Helvetica-Bold', fontSize=10, textColor=DARK_GRAY,
        leading=15)
    S['table_hdr'] = ParagraphStyle('table_hdr',
        fontName='Helvetica-Bold', fontSize=8, textColor=WHITE,
        alignment=TA_LEFT, leading=11, tracking=1)
    S['table_cell'] = ParagraphStyle('table_cell',
        fontName='Helvetica', fontSize=9, textColor=DARK_GRAY,
        leading=13, alignment=TA_LEFT)
    S['table_cell_r'] = ParagraphStyle('table_cell_r',
        fontName='Helvetica', fontSize=9, textColor=DARK_GRAY,
        leading=13, alignment=TA_RIGHT)
    S['table_cell_bold'] = ParagraphStyle('table_cell_bold',
        fontName='Helvetica-Bold', fontSize=9, textColor=BLACK,
        leading=13, alignment=TA_RIGHT)
    S['bullet'] = ParagraphStyle('bullet',
        fontName='Helvetica', fontSize=10, textColor=DARK_GRAY,
        leading=16, leftIndent=10, spaceAfter=4)
    S['footer'] = ParagraphStyle('footer',
        fontName='Helvetica', fontSize=7.5, textColor=LIGHT_GRAY,
        alignment=TA_CENTER, leading=10)
    S['confidential'] = ParagraphStyle('confidential',
        fontName='Helvetica', fontSize=7, textColor=LIGHT_GRAY,
        alignment=TA_CENTER, leading=10, tracking=2)
    S['positive'] = ParagraphStyle('positive',
        fontName='Helvetica-Bold', fontSize=9.5, textColor=GREEN_NUM,
        alignment=TA_RIGHT, leading=13)
    S['negative'] = ParagraphStyle('negative',
        fontName='Helvetica-Bold', fontSize=9.5, textColor=RED_NUM,
        alignment=TA_RIGHT, leading=13)
    S['neutral_r'] = ParagraphStyle('neutral_r',
        fontName='Helvetica', fontSize=9.5, textColor=MID_GRAY,
        alignment=TA_RIGHT, leading=13)
    return S


# ── Page callbacks ────────────────────────────────────────────────────────────
def on_cover_page(canvas, doc, data, agency, S):
    canvas.saveState()
    W_, H_ = A4

    # White background
    canvas.setFillColor(WHITE)
    canvas.rect(0, 0, W_, H_, fill=1, stroke=0)

    # Top navy band
    canvas.setFillColor(ACCENT)
    canvas.rect(0, H_ - 28*mm, W_, 28*mm, fill=1, stroke=0)

    # Gold line below band
    canvas.setFillColor(ACCENT2)
    canvas.rect(0, H_ - 29.5*mm, W_, 1.5*mm, fill=1, stroke=0)

    # Agency name in band
    canvas.setFont('Helvetica-Bold', 11)
    canvas.setFillColor(WHITE)
    agency_name = agency.get('agency_name', 'AGENCE IMMOBILIÈRE').upper()
    canvas.drawCentredString(W_/2, H_ - 13*mm, agency_name)
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#a0b4c8'))
    sub = []
    if agency.get('agency_address'): sub.append(agency['agency_address'])
    if agency.get('phone'): sub.append(agency['phone'])
    if sub:
        canvas.drawCentredString(W_/2, H_ - 20*mm, '   •   '.join(sub))

    # "AVIS DE VALEUR" label
    y = H_ - 50*mm
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(ACCENT2)
    canvas.drawCentredString(W_/2, y, 'A V I S   D E   V A L E U R')

    # Thin rule
    canvas.setStrokeColor(RULE_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(ML + 30*mm, y - 4*mm, W_ - MR - 30*mm, y - 4*mm)

    # Address
    addr = data.get('address', '')
    parts = [p.strip() for p in addr.split(',')]
    y_addr = y - 22*mm
    canvas.setFont('Helvetica-Bold', 17)
    canvas.setFillColor(BLACK)
    canvas.drawCentredString(W_/2, y_addr, parts[0] if parts else addr)
    if len(parts) > 1:
        canvas.setFont('Helvetica', 12)
        canvas.setFillColor(MID_GRAY)
        canvas.drawCentredString(W_/2, y_addr - 9*mm, ', '.join(parts[1:]))

    # Property badge
    prop = f"{data.get('propertyType','Appartement')}  ·  {data.get('surfaceArea','')} m²  ·  {data.get('rooms','')} pièces"
    canvas.setFont('Helvetica', 9)
    canvas.setFillColor(LIGHT_GRAY)
    canvas.drawCentredString(W_/2, y_addr - 17*mm, prop.upper())

    # Divider
    y_div = y_addr - 24*mm
    canvas.setStrokeColor(RULE_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(ML + 20*mm, y_div, W_ - MR - 20*mm, y_div)

    # Owners
    owners = data.get('owners', [])
    y_own = y_div - 12*mm
    canvas.setFont('Helvetica', 9)
    canvas.setFillColor(MID_GRAY)
    canvas.drawCentredString(W_/2, y_own, 'Préparé à l\'attention de')
    canvas.setFont('Helvetica-Bold', 13)
    canvas.setFillColor(BLACK)
    own_str = '  &  '.join(owners) if owners else 'Madame, Monsieur'
    canvas.drawCentredString(W_/2, y_own - 8*mm, own_str)

    # Price box — centered
    net = data.get('net_vendeur', 0)
    fai = data.get('fai_price', 0)
    y_box_top = y_own - 30*mm
    box_w = 90*mm; box_h = 38*mm
    box_x = (W_ - box_w) / 2

    # Light bg rect
    canvas.setFillColor(BG_CARD)
    canvas.roundRect(box_x, y_box_top - box_h, box_w, box_h, 4, fill=1, stroke=0)
    # Gold top border
    canvas.setFillColor(ACCENT2)
    canvas.roundRect(box_x, y_box_top - 1.5*mm, box_w, 1.5*mm, 1, fill=1, stroke=0)

    # Price
    canvas.setFont('Helvetica-Bold', 28)
    canvas.setFillColor(ACCENT)
    canvas.drawCentredString(W_/2, y_box_top - box_h/2 - 2*mm, fmt_price(net))

    canvas.setFont('Helvetica', 7.5)
    canvas.setFillColor(LIGHT_GRAY)
    canvas.drawCentredString(W_/2, y_box_top - box_h/2 - 9*mm, 'PRIX NET VENDEUR ESTIMÉ')

    if fai:
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(MID_GRAY)
        canvas.drawCentredString(W_/2, y_box_top - box_h + 5*mm, f"Prix affiché (honoraires inclus) : {fmt_price(fai)}")

    # Fourchette
    rmin = data.get('range_min',0); rmax = data.get('range_max',0)
    if rmin and rmax:
        y_range = y_box_top - box_h - 8*mm
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(LIGHT_GRAY)
        canvas.drawCentredString(W_/2, y_range, f"Fourchette d'estimation  :  {fmt_price(rmin)}  —  {fmt_price(rmax)}")

    # Bottom band
    canvas.setFillColor(ACCENT)
    canvas.rect(0, 0, W_, 18*mm, fill=1, stroke=0)
    canvas.setFillColor(ACCENT2)
    canvas.rect(0, 18*mm, W_, 1*mm, fill=1, stroke=0)

    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.HexColor('#a0b4c8'))
    agent = agency.get('agent_name','')
    canvas.drawString(ML, 11*mm, f"Conseiller : {agent}" if agent else '')
    canvas.drawCentredString(W_/2, 11*mm, date_fr())
    canvas.setFillColor(colors.HexColor('#a0b4c8'))
    canvas.drawRightString(W_ - MR, 11*mm, 'Document confidentiel')

    canvas.restoreState()


def on_normal_page(canvas, doc, agency, S):
    canvas.saveState()
    W_, H_ = A4

    # White bg
    canvas.setFillColor(WHITE)
    canvas.rect(0, 0, W_, H_, fill=1, stroke=0)

    # Top thin accent bar
    canvas.setFillColor(ACCENT)
    canvas.rect(0, H_ - 8*mm, W_, 8*mm, fill=1, stroke=0)
    canvas.setFillColor(ACCENT2)
    canvas.rect(0, H_ - 9*mm, W_, 1*mm, fill=1, stroke=0)

    # Agency name top right
    canvas.setFont('Helvetica-Bold', 7.5)
    canvas.setFillColor(WHITE)
    canvas.drawRightString(W_ - MR, H_ - 5.5*mm, agency.get('agency_name','VALORIA').upper())

    # Footer line
    canvas.setStrokeColor(RULE_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(ML, MB + 5*mm, W_ - MR, MB + 5*mm)

    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(LIGHT_GRAY)
    canvas.drawString(ML, MB + 2*mm, 'AVIS DE VALEUR  —  Document strictement confidentiel')
    canvas.drawRightString(W_ - MR, MB + 2*mm, str(doc.page))

    canvas.restoreState()


# ── Section header helper ─────────────────────────────────────────────────────
def section_hdr(num, title, S, story):
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(num, S['section_num']))
    story.append(Paragraph(title, S['section_title']))
    story.append(HRFlowable(width=CW, thickness=0.5, color=ACCENT2,
                            spaceAfter=5*mm))


# ── Page 2 : Lettre d'accompagnement ─────────────────────────────────────────
def build_letter(story, data, S, agency):
    story.append(Spacer(1, 8*mm))

    owners = data.get('owners', [])
    own_str = ' et '.join(owners) if owners else 'Madame, Monsieur'
    addr = data.get('address', '')
    prop_type = data.get('propertyType', 'votre bien').lower()
    surface = data.get('surfaceArea', '')
    rooms = data.get('rooms', '')
    agent = agency.get('agent_name', 'Votre conseiller')
    city = agency.get('city', 'le') or 'le'

    # Lieu & date
    story.append(Paragraph(f"{city.capitalize()}, le {date_fr()}", S['sign']))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(f"{own_str},", S['letter_bold']))
    story.append(Spacer(1, 16*mm))

    features = data.get('keyFeatures', [])
    feat_str = (', '.join(features[:2]) + ('…' if len(features) > 2 else '')) if features else 'ses caractéristiques'

    body1 = (
        f"Suite à notre rencontre et à la visite approfondie de votre {prop_type} "
        f"situé au <b>{addr}</b>, nous avons le plaisir de vous remettre "
        f"notre avis de valeur expert."
    )
    body2 = (
        "Ce document a été établi selon une méthodologie rigoureuse, fondée sur "
        "l'analyse des transactions immobilières récentes enregistrées par les notaires "
        "dans votre secteur (base Demandes de Valeurs Foncières), croisée avec une "
        "évaluation précise des atouts propres à votre bien."
    )
    body3 = (
        f"Votre bien présente des caractéristiques remarquables — "
        f"{'notamment ' + feat_str if feat_str else 'que nous avons soigneusement évaluées'} "
        f"— qui ont été intégralement prises en compte dans notre estimation."
    )
    body4 = (
        "Nous restons naturellement à votre disposition pour vous présenter notre "
        "plan de commercialisation personnalisé et vous accompagner à chaque étape "
        "de votre projet de vente dans les meilleures conditions."
    )

    for t in [body1, body2, body3, body4]:
        story.append(Paragraph(t, S['letter']))
        story.append(Spacer(1, 2*mm))

    story.append(Spacer(1, 8*mm))
    story.append(Paragraph("Bien cordialement,", S['sign']))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>{agent}</b>", S['sign_bold']))
    if agency.get('agency_name'):
        story.append(Paragraph(agency['agency_name'], S['sign']))
    if agency.get('phone'):
        story.append(Paragraph(agency['phone'], S['sign']))
    if agency.get('website'):
        story.append(Paragraph(agency['website'], S['sign']))


# ── Page 3 : Estimation & marché ─────────────────────────────────────────────
def build_estimation(story, data, S):
    story.append(PageBreak())
    section_hdr('01', 'Notre avis de valeur', S, story)

    net = data.get('net_vendeur', 0)
    fai = data.get('fai_price', 0)
    rmin = data.get('range_min', 0)
    rmax = data.get('range_max', 0)

    # Main price block
    price_data = [
        [Paragraph('PRIX NET VENDEUR ESTIMÉ', S['price_label'])],
        [Paragraph(fmt_price(net), S['price_big'])],
    ]
    if fai:
        price_data.append([Paragraph(f"Prix affiché (honoraires inclus) : <b>{fmt_price(fai)}</b>",
                                      S['body_center'])])

    pt = Table(price_data, colWidths=[CW])
    pt.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), BG_CARD),
        ('BOX', (0,0),(-1,-1), 0.5, ACCENT2),
        ('TOPPADDING', (0,0),(-1,-1), 10),
        ('BOTTOMPADDING', (0,0),(-1,-1), 10),
        ('ALIGN', (0,0),(-1,-1), 'CENTER'),
    ]))
    story.append(pt)
    story.append(Spacer(1, 4*mm))

    # Min / Max
    if rmin and rmax:
        rdata = [
            [Paragraph('FOURCHETTE BASSE', S['price_label']),
             Paragraph('', S['body']),
             Paragraph('FOURCHETTE HAUTE', S['price_label'])],
            [Paragraph(fmt_price(rmin), S['price_range']),
             Paragraph('—', S['body_center']),
             Paragraph(fmt_price(rmax), S['price_range'])],
        ]
        cw = [CW*0.4, CW*0.2, CW*0.4]
        rt = Table(rdata, colWidths=cw)
        rt.setStyle(TableStyle([
            ('ALIGN', (0,0),(-1,-1), 'CENTER'),
            ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0),(-1,-1), 6),
            ('BOTTOMPADDING', (0,0),(-1,-1), 6),
            ('LINEBELOW', (0,0),(-1,0), 0.3, RULE_GRAY),
        ]))
        story.append(rt)
        story.append(Spacer(1, 6*mm))

    # Market context
    story.append(Paragraph('Contexte du marché local', S['h2']))
    market = data.get('marketTrends', {})
    days = market.get('daysToSell', {}).get('apartment', 0)
    evol = market.get('apartmentEvolution', {}).get('oneYear', 0)
    rates = market.get('mortgageRates', {})
    r20 = rates.get('twentyYears', 0)

    ctx = (
        f"Le marché immobilier de votre secteur présente une dynamique "
        f"{'favorable' if evol >= 0 else 'en ajustement'}"
        f"{f', avec une évolution des prix de {evol:+.1f}% sur les douze derniers mois' if evol else ''}. "
        f"{'Le délai moyen de vente pour ce type de bien est actuellement de ' + str(days) + ' jours. ' if days else ''}"
        f"{'Les taux d\'emprunt à 20 ans s\'établissent autour de ' + str(r20) + '%, permettant de financer des projets immobiliers dans des conditions raisonnables.' if r20 else ''}"
    )
    story.append(Paragraph(ctx, S['body']))


# ── Page 4 : Construction du prix ────────────────────────────────────────────
def build_price_construction(story, data, S):
    story.append(PageBreak())
    section_hdr('02', 'Comment avons-nous établi ce prix ?', S, story)

    intro = (
        "Notre estimation repose sur une valeur de marché de référence, établie à partir "
        "des transactions notariales récentes comparables à votre bien, à laquelle nous "
        "appliquons des ajustements précis et documentés reflétant les spécificités de votre bien."
    )
    story.append(Paragraph(intro, S['body']))
    story.append(Spacer(1, 4*mm))

    # Base DVF
    dvf = data.get('dvfBasePrice', 0)
    if dvf:
        base_rows = [[
            Paragraph('Valeur de marché de référence (base DVF notaires)', S['h3']),
            Paragraph(fmt_price(dvf), S['table_cell_bold'])
        ]]
        bt = Table(base_rows, colWidths=[CW*0.72, CW*0.28])
        bt.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), BG_CARD),
            ('LINEABOVE', (0,0),(-1,0), 1, ACCENT),
            ('LINEBELOW', (0,0),(-1,-1), 0.5, RULE_GRAY),
            ('TOPPADDING', (0,0),(-1,-1), 10),
            ('BOTTOMPADDING', (0,0),(-1,-1), 10),
            ('LEFTPADDING', (0,0),(0,-1), 10),
            ('RIGHTPADDING', (1,0),(1,-1), 10),
            ('ALIGN', (1,0),(1,-1), 'RIGHT'),
            ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
        ]))
        story.append(bt)
        story.append(Spacer(1, 5*mm))

    # Adjustments
    adjs = data.get('adjustmentsBreakdown', [])
    if adjs:
        story.append(Paragraph('Ajustements qualitatifs', S['h2']))

        hdr = [
            Paragraph('Critère', S['table_hdr']),
            Paragraph('Description', S['table_hdr']),
            Paragraph('Impact', S['table_hdr']),
        ]
        rows = [hdr]
        for adj in adjs:
            cat = adj.get('category', 'neutral')
            val = adj.get('value', 0)
            amt = adj.get('amountEur', 0)

            if amt:
                val_str = fmt_price(amt)
            elif val:
                val_str = f"+{val}%" if cat == 'positive' else f"{val}%"
            else:
                val_str = '—'

            if cat == 'positive':
                val_style = S['positive']
            elif cat == 'negative':
                val_style = S['negative']
            else:
                val_style = S['neutral_r']

            rows.append([
                Paragraph(adj.get('label', ''), S['table_cell']),
                Paragraph(adj.get('description', ''), S['body_sm']),
                Paragraph(val_str, val_style),
            ])

        at = Table(rows, colWidths=[CW*0.28, CW*0.52, CW*0.20])
        row_bg = [WHITE, BG_LIGHT]
        ts = [
            ('BACKGROUND', (0,0),(-1,0), ACCENT),
            ('ROWBACKGROUNDS', (0,1),(-1,-1), row_bg),
            ('BOX', (0,0),(-1,-1), 0.3, RULE_GRAY),
            ('LINEBELOW', (0,0),(-1,-2), 0.3, RULE_GRAY),
            ('ALIGN', (2,0),(2,-1), 'RIGHT'),
            ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0),(-1,-1), 7),
            ('BOTTOMPADDING', (0,0),(-1,-1), 7),
            ('LEFTPADDING', (0,0),(0,-1), 8),
            ('LEFTPADDING', (1,0),(1,-1), 8),
            ('RIGHTPADDING', (2,0),(2,-1), 8),
        ]
        at.setStyle(TableStyle(ts))
        story.append(at)


# ── Page 5 : Comparables DVF ─────────────────────────────────────────────────
def build_dvf(story, data, S):
    story.append(PageBreak())
    section_hdr('03', 'Références du marché', S, story)

    intro = (
        "Les transactions ci-dessous ont été réalisées dans votre secteur géographique "
        "et enregistrées par les notaires. Elles constituent la base objective de notre "
        "analyse et permettent d'ancrer notre estimation dans la réalité du marché."
    )
    story.append(Paragraph(intro, S['body']))
    story.append(Spacer(1, 4*mm))

    comps = data.get('comparableSales', [])
    if not comps:
        story.append(Paragraph("Aucune transaction comparable disponible.", S['body_sm']))
        return

    hdr = [
        Paragraph('Adresse', S['table_hdr']),
        Paragraph('Date', S['table_hdr']),
        Paragraph('Surface', S['table_hdr']),
        Paragraph('Prix de vente', S['table_hdr']),
        Paragraph('Prix / m²', S['table_hdr']),
    ]
    rows = [hdr]
    for c in comps[:8]:
        surf = c.get('surface', 0)
        price = c.get('price', 0)
        ppm2 = int(price / surf) if surf else 0
        rows.append([
            Paragraph(c.get('address','')[:38], S['table_cell']),
            Paragraph(c.get('date','')[:7], S['table_cell']),
            Paragraph(f"{surf} m²" if surf else '—', S['table_cell']),
            Paragraph(fmt_price(price), S['table_cell_r']),
            Paragraph(f"{ppm2:,}".replace(',',' ') + ' €/m²' if ppm2 else '—', S['table_cell_r']),
        ])

    cw = [CW*0.36, CW*0.11, CW*0.11, CW*0.22, CW*0.20]
    ct = Table(rows, colWidths=cw)
    ct.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,0), ACCENT),
        ('ROWBACKGROUNDS', (0,1),(-1,-1), [WHITE, BG_LIGHT]),
        ('BOX', (0,0),(-1,-1), 0.3, RULE_GRAY),
        ('LINEBELOW', (0,0),(-1,-2), 0.3, RULE_GRAY),
        ('ALIGN', (1,0),(-1,-1), 'CENTER'),
        ('ALIGN', (3,0),(-1,-1), 'RIGHT'),
        ('ALIGN', (4,0),(4,-1), 'RIGHT'),
        ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0),(-1,-1), 7),
        ('BOTTOMPADDING', (0,0),(-1,-1), 7),
        ('LEFTPADDING', (0,0),(0,-1), 8),
        ('RIGHTPADDING', (4,0),(4,-1), 8),
    ]))
    story.append(ct)


# ── Page 6 : Votre bien ───────────────────────────────────────────────────────
def build_property(story, data, S):
    story.append(PageBreak())
    section_hdr('04', 'Les atouts de votre bien', S, story)

    features = data.get('keyFeatures', [])
    if features:
        rows = []
        for f in features:
            rows.append([
                Paragraph('—', S['body_sm']),
                Paragraph(f, S['body'])
            ])
        ft = Table(rows, colWidths=[5*mm, CW - 5*mm])
        ft.setStyle(TableStyle([
            ('ALIGN', (0,0),(0,-1), 'CENTER'),
            ('VALIGN', (0,0),(-1,-1), 'TOP'),
            ('TOPPADDING', (0,0),(-1,-1), 3),
            ('BOTTOMPADDING', (0,0),(-1,-1), 3),
            ('LEFTPADDING', (1,0),(1,-1), 6),
        ]))
        story.append(ft)
        story.append(Spacer(1, 5*mm))

    # DPE
    dpe = data.get('dpeRating', '')
    if dpe:
        dpe_colors = {'A':'#00cc44','B':'#50cd3a','C':'#a0c020','D':'#e0c000',
                      'E':'#e08000','F':'#e04000','G':'#c00000'}
        c = dpe_colors.get(dpe.upper(), '#999')
        story.append(Paragraph(
            f"Diagnostic de Performance Énergétique (DPE) : "
            f"<font color='{c}'><b>Classe {dpe.upper()}</b></font>",
            S['body']))


# ── Page 7 : Analyse patrimoniale ────────────────────────────────────────────
def build_patrimonial(story, data, S):
    purchase_price = data.get('purchasePrice', 0)
    net = data.get('net_vendeur', 0)
    if not purchase_price:
        return

    story.append(PageBreak())
    section_hdr('05', 'Analyse patrimoniale', S, story)

    delta = net - purchase_price
    delta_pct = (delta / purchase_price * 100) if purchase_price else 0
    sign = '+' if delta >= 0 else ''
    delta_color = '#2d6a4f' if delta >= 0 else '#9b2226'

    hist = [
        [Paragraph('Prix d\'acquisition', S['h3']),
         Paragraph(fmt_price(purchase_price), S['table_cell_bold'])],
        [Paragraph('Estimation actuelle', S['h3']),
         Paragraph(fmt_price(net), S['table_cell_bold'])],
        [Paragraph('Évolution de valeur', S['h3']),
         Paragraph(f"<font color='{delta_color}'><b>{sign}{fmt_price(delta)} ({sign}{delta_pct:.1f}%)</b></font>",
                   S['table_cell_bold'])],
    ]
    ht = Table(hist, colWidths=[CW*0.6, CW*0.4])
    ht.setStyle(TableStyle([
        ('ROWBACKGROUNDS', (0,0),(-1,-1), [WHITE, BG_LIGHT, WHITE]),
        ('BOX', (0,0),(-1,-1), 0.3, RULE_GRAY),
        ('LINEBELOW', (0,0),(-1,-2), 0.3, RULE_GRAY),
        ('TOPPADDING', (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('LEFTPADDING', (0,0),(0,-1), 10),
        ('ALIGN', (1,0),(1,-1), 'RIGHT'),
        ('RIGHTPADDING', (1,0),(1,-1), 10),
        ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
    ]))
    story.append(ht)
    story.append(Spacer(1, 6*mm))

    # Pinel alert if applicable
    reasoning = data.get('reasoning', '')
    if 'pinel' in (reasoning or '').lower():
        story.append(Paragraph('Point d\'attention fiscal — Dispositif Pinel', S['h2']))
        pinel_text = (
            "Votre bien ayant été acquis sous le dispositif Pinel, nous attirons votre attention "
            "sur un point important : ce dispositif impose un engagement de location sur une durée "
            "minimale. Une vente avant l'échéance de cet engagement entraînerait le remboursement "
            "intégral de toutes les réductions d'impôt perçues depuis l'acquisition, augmentées "
            "des intérêts de retard. Nous vous recommandons vivement de vérifier la date exacte "
            "de fin de votre engagement auprès de votre notaire ou de votre conseiller fiscal "
            "avant toute décision."
        )
        # Alert box
        alert_rows = [[Paragraph(pinel_text, S['body'])]]
        alrt = Table(alert_rows, colWidths=[CW])
        alrt.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), colors.HexColor('#fef9ec')),
            ('LINEABOVE', (0,0),(-1,0), 2, ACCENT2),
            ('BOX', (0,0),(-1,-1), 0.5, ACCENT2),
            ('TOPPADDING', (0,0),(-1,-1), 10),
            ('BOTTOMPADDING', (0,0),(-1,-1), 10),
            ('LEFTPADDING', (0,0),(-1,-1), 12),
            ('RIGHTPADDING', (0,0),(-1,-1), 12),
        ]))
        story.append(alrt)


# ── Page 8 : Stratégie de commercialisation ───────────────────────────────────
def build_strategy(story, data, S):
    strategy = data.get('agentStrategy', {})
    if not strategy:
        return

    story.append(PageBreak())
    section_hdr('06', 'Notre stratégie pour votre bien', S, story)

    angle = strategy.get('marketingAngle', '')
    if angle:
        story.append(Paragraph('Positionnement & cible acquéreur', S['h2']))
        story.append(Paragraph(angle, S['body']))
        story.append(Spacer(1, 4*mm))

    positioning = strategy.get('pricePositioning', '')
    if positioning:
        story.append(Paragraph('Stratégie de prix', S['h2']))
        pos_rows = [[Paragraph(positioning, S['body'])]]
        pt = Table(pos_rows, colWidths=[CW])
        pt.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), BG_CARD),
            ('LINEABOVE', (0,0),(-1,0), 1.5, ACCENT),
            ('BOX', (0,0),(-1,-1), 0.3, RULE_GRAY),
            ('TOPPADDING', (0,0),(-1,-1), 10),
            ('BOTTOMPADDING', (0,0),(-1,-1), 10),
            ('LEFTPADDING', (0,0),(-1,-1), 12),
            ('RIGHTPADDING', (0,0),(-1,-1), 12),
        ]))
        story.append(pt)
        story.append(Spacer(1, 5*mm))

    objections = strategy.get('objectionHandling', [])
    if objections:
        story.append(Paragraph('Points clés à anticiper', S['h2']))
        story.append(Paragraph(
            "Voici les points que nous avons identifiés et pour lesquels nous avons "
            "préparé des éléments de réponse précis à destination des acquéreurs :",
            S['body']))
        story.append(Spacer(1, 3*mm))

        for i, obj in enumerate(objections, 1):
            rows = [[
                Paragraph(f"{i:02d}", S['h3']),
                Paragraph(obj, S['body'])
            ]]
            ot = Table(rows, colWidths=[8*mm, CW - 8*mm])
            ot.setStyle(TableStyle([
                ('BACKGROUND', (0,0),(-1,-1), WHITE),
                ('BACKGROUND', (0,0),(0,0), BG_CARD),
                ('BOX', (0,0),(-1,-1), 0.3, RULE_GRAY),
                ('LINEBELOW', (0,0),(-1,-1), 0.3, RULE_GRAY),
                ('ALIGN', (0,0),(0,0), 'CENTER'),
                ('VALIGN', (0,0),(-1,-1), 'TOP'),
                ('TOPPADDING', (0,0),(-1,-1), 8),
                ('BOTTOMPADDING', (0,0),(-1,-1), 8),
                ('LEFTPADDING', (0,0),(0,0), 4),
                ('LEFTPADDING', (1,0),(1,0), 10),
            ]))
            story.append(ot)
            story.append(Spacer(1, 2*mm))


# ── Dernière page : mention légale ───────────────────────────────────────────
def build_legal(story, S, agency):
    story.append(PageBreak())
    story.append(Spacer(1, 20*mm))
    story.append(HRFlowable(width=CW, thickness=0.5, color=RULE_GRAY, spaceAfter=6*mm))

    legal = (
        "Cet avis de valeur a été établi à titre indicatif par "
        f"{agency.get('agency_name','notre agence')}, sur la base des informations "
        "disponibles à la date d'établissement du présent document et des données "
        "de transactions enregistrées par les notaires (source DVF). "
        "Il ne constitue pas une expertise immobilière au sens juridique du terme "
        "et ne saurait engager la responsabilité de l'agence au-delà de son rôle "
        "de conseil. La valeur finale de cession dépendra des conditions du marché "
        "au moment de la mise en vente, de la négociation entre les parties et de "
        "l'issue des éventuelles expertises techniques (diagnostics, financement)."
    )
    story.append(Paragraph(legal, S['body_sm']))
    story.append(Spacer(1, 4*mm))

    if agency.get('website') or agency.get('phone'):
        contact = '   ·   '.join(filter(None, [
            agency.get('agency_name',''),
            agency.get('phone',''),
            agency.get('website',''),
        ]))
        story.append(Paragraph(contact, S['footer']))


# ── Main generator ────────────────────────────────────────────────────────────
def generate_owner_pdf(report_data, agency_data, output):
    """output : chemin fichier (str) ou buffer BytesIO"""
    S = make_styles()

    class ValoriaDoc(BaseDocTemplate):
        def __init__(self, filename, **kwargs):
            super().__init__(filename, **kwargs)
            # Cover frame (full bleed — no margins, canvas handles all drawing)
            f_cover = Frame(0, 0, W, H, leftPadding=0, rightPadding=0,
                           topPadding=0, bottomPadding=0, id='cover')
            # Normal content frame
            f_normal = Frame(ML, MB + 12*mm, CW, H - MT - MB - 20*mm,
                            leftPadding=0, rightPadding=0,
                            topPadding=0, bottomPadding=0, id='normal')
            pt_cover = PageTemplate('cover', [f_cover],
                onPage=lambda c, d: on_cover_page(c, d, report_data, agency_data, S))
            pt_normal = PageTemplate('normal', [f_normal],
                onPage=lambda c, d: on_normal_page(c, d, agency_data, S))
            self.addPageTemplates([pt_cover, pt_normal])

    doc = ValoriaDoc(output, pagesize=A4,
                     leftMargin=ML, rightMargin=MR,
                     topMargin=MT, bottomMargin=MB + 12*mm)

    story = []

    # Page 1 : cover — template cover, contenu dessiné entièrement par onPage
    # Un Spacer minuscule suffit à "occuper" le frame vide sans déborder
    from reportlab.platypus import Spacer as _Sp
    story.append(NextPageTemplate('cover'))
    story.append(_Sp(1, 1))   # occupe le frame sans contenu visible
    story.append(NextPageTemplate('normal'))
    story.append(PageBreak())  # force le saut vers le template normal

    # Page 2 : lettre
    build_letter(story, report_data, S, agency_data)

    # Page 3 : estimation
    build_estimation(story, report_data, S)

    # Page 4 : construction du prix
    build_price_construction(story, report_data, S)

    # Page 5 : DVF
    build_dvf(story, report_data, S)

    # Page 6 : bien
    build_property(story, report_data, S)

    # Page 7 : patrimonial (si prix d'achat connu)
    build_patrimonial(story, report_data, S)

    # Page 8 : stratégie
    build_strategy(story, report_data, S)

    # Dernière page : mention légale
    build_legal(story, S, agency_data)

    doc.build(story)
    return output


# ── Test ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    report = {
        "address": "35 Avenue Charles de Gaulle, 84130 Le Pontet",
        "propertyType": "Appartement", "surfaceArea": "89", "rooms": 4,
        "dpeRating": "C",
        "net_vendeur": 191676, "fai_price": 201210,
        "range_min": 173710, "range_max": 209642,
        "owners": ["M. NOVAIN Florent", "Mme GAUMON Stéphanie"],
        "purchasePrice": 289000,
        "keyFeatures": [
            "Appartement 4 pièces de 89 m² en parfait état, livré fin 2021",
            "Jardin privatif d'environ 50 m² et terrasse",
            "Double stationnement : garage fermé (Lot 20) et place extérieure (Lot 33)",
            "Charges de copropriété très faibles — environ 103 €/mois",
            "Taxe foncière 2025 de 1 337 €",
            "Actuellement loué 995 € hors charges sous dispositif Pinel",
        ],
        "dvfBasePrice": 140798,
        "adjustmentsBreakdown": [
            {"category":"positive","label":"Vue dégagée","value":3,"description":"Luminosité et qualité de vie optimales."},
            {"category":"positive","label":"Espace extérieur","value":8,"description":"Jardin privatif 50 m² et terrasse, atout rare pour un appartement."},
            {"category":"positive","label":"État général","value":3,"description":"Construction très récente, aucuns travaux à prévoir."},
            {"category":"positive","label":"Garage fermé","value":0,"amountEur":15000,"description":"Valeur intrinsèque du garage fermé (Lot 20) sur le marché local."},
            {"category":"positive","label":"Bonus stationnement","value":4.5,"description":"Prime de commodité liée à la présence d'un garage fermé dans ce secteur."},
            {"category":"positive","label":"Parking extérieur","value":0,"amountEur":6000,"description":"Place de stationnement extérieure supplémentaire (Lot 33)."},
        ],
        "comparableSales": [
            {"address":"48 BD EMILE ZOLA, 84130 Le Pontet","date":"2023-11","surface":72,"price":126750},
            {"address":"9090 BD EMILE ZOLA, 84130 Le Pontet","date":"2023-09","surface":72,"price":93000},
            {"address":"26 AV CHARLES DE GAULLE, 84130 Le Pontet","date":"2023-07","surface":89,"price":165800},
            {"address":"26 AV CHARLES DE GAULLE, 84130 Le Pontet","date":"2023-11","surface":89,"price":182490},
            {"address":"26 AV CHARLES DE GAULLE, 84130 Le Pontet","date":"2023-03","surface":88,"price":144770},
            {"address":"4 IMP DU FELIBRIGE, 84130 Le Pontet","date":"2024-11","surface":80,"price":177450},
        ],
        "marketTrends": {
            "apartmentEvolution": {"oneYear": 5.1},
            "daysToSell": {"apartment": 115},
            "mortgageRates": {"twentyYears": 3.21}
        },
        "agentStrategy": {
            "marketingAngle": "Valoriser la rareté d'un appartement 4 pièces récent avec grand jardin de 50 m², double stationnement et charges de copropriété extrêmement faibles.",
            "pricePositioning": "Positionnement en fourchette haute justifié par les annexes exceptionnelles (jardin, double stationnement) et la parfaite maîtrise des charges. Cible : investisseurs ou acquéreurs en recherche de confort.",
            "objectionHandling": [
                "Dispositif Pinel : la vente avant la fin de l'engagement locatif implique le remboursement des avantages fiscaux perçus. Nous vous conseillons de vérifier la date d'échéance avec votre notaire.",
                "Incohérence d'adresse (37 vs 35 avenue) : les documents administratifs portent le numéro 37, à faire vérifier et rectifier avant la signature du mandat.",
                "Régularisation locataire : un solde de 531 € est en cours d'apurement — ce point ne remet pas en cause la solidité du dossier.",
            ]
        },
        "reasoning": "pinel engagement non terminé",
    }
    agency = {
        "agent_name": "Jean-Pierre Martin",
        "agency_name": "Idéal Home Avignon",
        "agency_address": "12 Rue de la République, 84000 Avignon",
        "phone": "04 90 00 00 00",
        "website": "www.idealhome-avignon.fr",
        "city": "Avignon",
    }
    out = generate_owner_pdf(report, agency, '/home/claude/owner_report_white.pdf')
    print(f"✓ PDF généré : {out}")
