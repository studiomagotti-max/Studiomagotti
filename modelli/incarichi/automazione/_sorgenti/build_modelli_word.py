"""Genera i modelli Word con campi {{...}} per l'automazione di invio mandati.
Sorgenti: modelli/incarichi/MOD_lettera_incarico_preventivo.docx e
modelli/antiriciclaggio/MOD_AML_modulistica.docx. Eseguire dalla root del repository."""
import copy, sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

OUT = 'modelli/incarichi/automazione/'
W = qn

def base_rpr(p):
    r = p.runs[0] if p.runs else None
    if r is None or r._r.rPr is None:
        return None
    rpr = copy.deepcopy(r._r.rPr)
    for tag in ('w:highlight', 'w:highlightCs', 'w:b', 'w:bCs'):
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)
    return rpr

def set_runs(p, segs, rpr=None):
    """segs: lista di (testo, tipo) con tipo 'n' normale, 'f' campo (giallo), 'b' grassetto, 'm' marcatore."""
    if rpr is None:
        rpr = base_rpr(p)
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for text, kind in segs:
        run = p.add_run(text)
        if rpr is not None:
            run._r.insert(0, copy.deepcopy(rpr))
        if kind in ('f', 'm'):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        if kind == 'b':
            run.bold = True
    return p

def F(k):
    return ('{{%s}}' % k, 'f')

def M(k):
    return ('{{?%s}}' % k, 'm')

def N(t):
    return (t, 'n')

def insert_paragraph_after(p, segs, bold=False):
    new_p = copy.deepcopy(p._p)
    p._p.addnext(new_p)
    np_ = Paragraph(new_p, p._parent)
    set_runs(np_, segs)
    if bold:
        for r in np_.runs:
            r.bold = True
    return np_

def clone_row_after(table, row_idx):
    tr = table.rows[row_idx]._tr
    new = copy.deepcopy(tr)
    tr.addnext(new)
    return new

# ---------------------------------------------------------------- MANDATO
def build_mandato():
    d = Document('modelli/incarichi/MOD_lettera_incarico_preventivo.docx')
    body = d.element.body
    # rimuove riquadro istruzioni (primo elemento) e paragrafo vuoto seguente
    first = list(body.iterchildren())[0]
    body.remove(first)
    P = d.paragraphs
    # indici riferiti al documento originale dopo la rimozione della tabella: invariati per i paragrafi
    set_runs(P[11], [N('Mantova, '), F('DATA_LETTERA')])
    set_runs(P[13], [F('DENOMINAZIONE')])
    for r in P[13].runs: r.bold = True
    set_runs(P[14], [F('INDIRIZZO_COMPLETO')])
    set_runs(P[15], [N('C.F. '), F('CF'), N(' – P.IVA '), F('PIVA')])
    set_runs(P[16], [N('PEC: '), F('PEC_CLIENTE')])
    t18 = [r.text for r in P[18].runs]
    set_runs(P[18], [N(t18[0]), F('DENOMINAZIONE'), F('FRASE_LEGALE_RAPP'), N(t18[4])])
    set_runs(P[21], [N("Il Cliente conferisce al Professionista l'incarico di svolgere le seguenti prestazioni:")])
    set_runs(P[22], [M('P_CONT'), N('Tenuta della contabilità '), F('REGIME_CONTABILE'),
                     N(' e dei registri IVA; liquidazioni periodiche IVA e comunicazioni LIPE;')])
    set_runs(P[23], [M('P_BIL'), N("Redazione del progetto di bilancio d'esercizio in forma "), F('SCHEMA_BILANCIO'),
                     N(' e della nota integrativa'), F('TXT_RELAZIONE_GESTIONE'), N('; deposito presso il Registro Imprese;')])
    set_runs(P[24], [M('P_DICH'), N('Predisposizione e trasmissione telematica delle dichiarazioni fiscali: Redditi, IRAP, IVA'),
                     F('TXT_770'), N(';')])
    set_runs(P[25], [M('P_F24'), N('Calcolo delle imposte e predisposizione dei modelli F24 per i versamenti; adempimenti relativi agli ISA'),
                     F('TXT_CPB'), N(';')])
    set_runs(P[26], [M('P_SOC'), N(P[26].text.lstrip('☐ '))])
    set_runs(P[27], [M('P_CONS'), N(P[27].text.lstrip('☐ '))])
    set_runs(P[28], [M('P_STRA'), N("Assistenza nell'operazione straordinaria di "), F('DESCR_STRA'), N(", come descritta nell'Allegato A;")])
    set_runs(P[29], [M('P_PER'), N('Perizia di stima / consulenza tecnica di parte relativa a '), F('DESCR_PER'), N(';')])
    set_runs(P[30], [M('P_ALT'), N('Altro: '), F('DESCR_ALT'), N('.')])
    t = [r.text for r in P[40].runs]
    set_runs(P[40], [N(t[0]), F('GG_CONSEGNA_DOC'), N(t[2]), F('DATA_CONSEGNA_DICH'), N(t[4])])
    t = [r.text for r in P[45].runs]
    set_runs(P[45], [N(t[0]), F('COMPLESSITA'), N(t[2]), F('MOTIVAZIONE_COMPLESSITA'), N('.')])
    t = [r.text for r in P[46].runs]
    set_runs(P[46], [N(t[0]), F('TARIFFA_TITOLARE'), N(t[2]), F('TARIFFA_COLLABORATORI'), N(t[4])])
    set_runs(P[47], [N('Sono dovuti il rimborso delle spese sostenute per conto del Cliente (diritti, bolli, imposte, spese di spedizione), addebitate analiticamente'),
                     F('TXT_SPESE_GENERALI'), N('.')])
    t = [r.text for r in P[48].runs]
    set_runs(P[48], [N(t[0]), F('CASSA_PREVIDENZA'), N(t[2])])
    set_runs(P[49], [N('I compensi forfetari sono aggiornati annualmente '), F('AGGIORNAMENTO_COMPENSI'),
                     N(', con comunicazione scritta entro il '), F('DATA_COMUNICAZIONE_AGG'), N('.')])
    set_runs(P[50], [M('COMMITTENTE_FORTE'), N('Ai sensi della L. 21 aprile 2023, n. 49, il compenso pattuito è proporzionato alla quantità e qualità del lavoro svolto e al contenuto e alle caratteristiche della prestazione, e non è inferiore ai parametri del D.M. 2 luglio 2012, n. 140.')])
    set_runs(P[52], [F('TXT_ACCONTO')])
    set_runs(P[54], [N('Fatturazione: '), F('FATTURAZIONE'), N('.')])
    set_runs(P[55], [N('Pagamento entro '), F('GG_PAGAMENTO'), N(' giorni dal ricevimento della fattura mediante '), F('TXT_MODALITA_PAGAMENTO'), N('.')])
    set_runs(P[56], [M('NON_CONSUMATORE'), N(P[56].text)])
    t = [r.text for r in P[57].runs]
    set_runs(P[57], [N(t[0]), F('N_FATTURE_RECESSO'), N(t[2])])
    t = [r.text for r in P[63].runs]
    set_runs(P[63], [N(t[0]), F('DURATA'), N(t[2]), F('GG_DISDETTA'), N(t[4])])
    set_runs(P[71], [M('CLAUSOLA_ART28'), N("Per i dati personali di terzi trattati per conto del Cliente nell'ambito della tenuta della contabilità, il Professionista agisce quale responsabile del trattamento ai sensi dell'art. 28 Reg. UE 2016/679, secondo l'atto di nomina sottoscritto separatamente.")])
    set_runs(P[75], [N("Le comunicazioni relative all'incarico sono effettuate ai seguenti indirizzi: Professionista PEC maurizio.magotti@pec.it; Cliente "), F('RECAPITO_COMUNICAZIONI'), N('.')])
    set_runs(P[77], [N("Per le controversie relative al compenso le parti possono richiedere il tentativo di conciliazione presso l'Ordine dei Dottori Commercialisti e degli Esperti Contabili di Mantova. "), F('TXT_FORO')])
    set_runs(P[79], [N('Mantova, '), F('DATA_LETTERA')])
    set_runs(P[86], [N('Trasmessa unitamente alla presente in file separato (dichiarazione M3 e informativa M1).')])
    set_runs(P[88], [N('Trasmessa unitamente alla presente in file separato.')])
    # Allegato D – SEPA (condizionale): titolo con il formato del titolo Allegato C
    pD1 = Paragraph(copy.deepcopy(P[87]._p), P[87]._parent)
    P[88]._p.addnext(pD1._p)
    set_runs(pD1, [M('SDD'), N('ALLEGATO D – MANDATO PER ADDEBITO DIRETTO SEPA CORE (Reg. UE 260/2012)')],
             rpr=copy.deepcopy(P[87].runs[0]._r.rPr))
    pD2 = Paragraph(copy.deepcopy(P[88]._p), P[88]._parent)
    pD1._p.addnext(pD2._p)
    set_runs(pD2, [M('SDD'), N('Trasmesso unitamente alla presente in file separato.')])
    # Tabelle: [0]=polizza [1]=firme [2]=compensi (dopo rimozione riquadro)
    tp = d.tables[0]
    vals = ['RC_COMPAGNIA', 'RC_POLIZZA', 'RC_MASS_SINISTRO', 'RC_MASS_ANNUO', 'RC_SCADENZA', 'RC_SINDACO_REVISORE']
    for i, k in enumerate(vals, start=1):
        cp = tp.rows[i].cells[1].paragraphs[0]
        pre = [N('€ ')] if k.startswith('RC_MASS') else []
        set_runs(cp, pre + [F(k)])
    tf = d.tables[1]
    set_runs(tf.rows[1].cells[1].paragraphs[0], [F('DENOMINAZIONE'), F('TXT_FIRMA_LR')])
    tc = d.tables[2]
    # righe originali: 1 cont,2 bil,3 dich,4 soc,5 cons,6 stra,7 tempo,8 spese,9 totale
    # inserisce F24 dopo dich (riga 3), perizia e altro dopo stra (riga 6)
    def fill(row, flag, label_segs, key):
        cells = row.cells
        set_runs(cells[0].paragraphs[0], [M(flag)] + label_segs)
        set_runs(cells[1].paragraphs[0], [F('MOD_' + key)])
        set_runs(cells[2].paragraphs[0], [F('IMP_' + key)])
        set_runs(cells[3].paragraphs[0], [F('PER_' + key)])
    clone_row_after(tc, 6); clone_row_after(tc, 6)  # righe 7,8 nuove (perizia, altro)
    clone_row_after(tc, 3)  # riga 4 nuova (F24)
    R = tc.rows
    fill(R[1], 'P_CONT', [N('Tenuta contabilità e adempimenti IVA')], 'CONT')
    fill(R[2], 'P_BIL', [N("Bilancio d'esercizio e deposito")], 'BIL')
    fill(R[3], 'P_DICH', [N('Dichiarazioni fiscali (Redditi, IRAP, IVA)')], 'DICH')
    fill(R[4], 'P_F24', [N('Calcolo imposte, F24, ISA'), F('TXT_CPB_TAB')], 'F24')
    fill(R[5], 'P_SOC', [N('Adempimenti societari ordinari')], 'SOC')
    fill(R[6], 'P_CONS', [N('Consulenza ordinaria continuativa')], 'CONS')
    fill(R[7], 'P_STRA', [N('Operazione straordinaria: '), F('DESCR_STRA')], 'STRA')
    fill(R[8], 'P_PER', [N('Perizia / consulenza tecnica: '), F('DESCR_PER')], 'PER')
    fill(R[9], 'P_ALT', [N('Altro: '), F('DESCR_ALT')], 'ALT')
    set_runs(R[10].cells[2].paragraphs[0], [F('TARIFFA_TITOLARE'), N(' /ora (titolare); '), F('TARIFFA_COLLABORATORI'), N(' /ora (collaboratori)')])
    set_runs(R[11].cells[0].paragraphs[0], [M('SPESE_GENERALI'), N('Rimborso forfetario spese generali')])
    set_runs(R[11].cells[1].paragraphs[0], [F('SPESE_GENERALI_PERC'), N('% del compenso')])
    set_runs(R[12].cells[0].paragraphs[0], [N('Totale annuo (prestazioni periodiche)')])
    from docx.oxml import OxmlElement
    for row in tc.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
    for r in R[12].cells[0].paragraphs[0].runs: r.bold = True
    set_runs(R[12].cells[2].paragraphs[0], [F('TOTALE_ANNUO')])
    for r in R[12].cells[2].paragraphs[0].runs: r.bold = True
    # piè di pagina
    for sec in d.sections:
        for fp in sec.footer.paragraphs:
            for r in fp.runs:
                if r.text == '[Cliente]':
                    r.text = '{{DENOMINAZIONE}}'
    d.save(OUT + 'MOD_mandato_campi.docx')

# ---------------------------------------------------------------- AML (M1 + M3)
def build_aml():
    d = Document('modelli/antiriciclaggio/MOD_AML_modulistica.docx')
    body = d.element.body
    els = list(body.iterchildren())
    keep = set(range(2, 18)) | set(range(47, 82))
    sect = els[-1] if els[-1].tag == qn('w:sectPr') else None
    for i, el in enumerate(els):
        if el is sect:
            continue
        if i not in keep:
            body.remove(el)
    blank = '____________________'
    def neutralize(p, mapping):
        for r in p.runs:
            if r.font.highlight_color:
                key = mapping.get(r.text)
                r.font.highlight_color = None
                if key:
                    r.text = key
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                else:
                    r.text = blank
    mp = {'[Cognome Nome]': '{{FIRMATARIO}}',
          '[cliente / legale rappresentante di DENOMINAZIONE, C.F. ______]': '{{QUALITA_DICHIARANTE}}',
          '[denominazione, sede, n. Registro Imprese e C.F.]': '{{DENOMINAZIONE}}, C.F. {{CF}}',
          '[denominazione]': '{{DENOMINAZIONE}}'}
    for p in d.paragraphs:
        if p.text.startswith('Mantova, [GG/MM/AAAA]') or p.text.startswith('Per presa visione – Mantova'):
            pre = 'Per presa visione – ' if p.text.startswith('Per presa') else ''
            set_runs(p, [N(pre + 'Luogo e data: ____________________')])
            continue
        neutralize(p, mp)
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    neutralize(p, mp)
    for sec in d.sections:
        # piè di pagina: sostituisce il testo descrittivo mantenendo i campi pagina
        for fp in sec.footer.paragraphs:
            for r in fp.runs:
                if r.text.startswith('Modulistica AML'):
                    r.text = 'Antiriciclaggio – {{DENOMINAZIONE}} – pag. '
                elif r.text in ('[1.0]', ' – pag. ') or r.font.highlight_color:
                    r.text = ''
                    r.font.highlight_color = None
    d.save(OUT + 'MOD_AML_cliente_campi.docx')

# ---------------------------------------------------------------- documenti nuovi (SEPA, privacy)
def new_doc():
    d = Document()
    sec = d.sections[0]
    from docx.shared import Cm
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = d.styles['Normal']
    st.font.name = 'Arial'; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    st.paragraph_format.space_after = Pt(4)
    return d

def para(d, segs, align=None, size=None, bold=False, space_after=None):
    p = d.add_paragraph()
    for text, kind in segs:
        r = p.add_run(text)
        if kind in ('f', 'm'):
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        if kind == 'b' or bold:
            r.bold = True
        if size:
            r.font.size = Pt(size)
    if align:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p

def intestazione(d):
    para(d, [N('STUDIO COMMERCIALISTA MAURIZIO MAGOTTI')], bold=True, size=11, space_after=0)
    para(d, [N('C.so Vittorio Emanuele II°, n. 101 – 46100 Mantova – Tel. +39 0376 360750 – info@studiomagotti.com – PEC maurizio.magotti@pec.it')], size=8, space_after=10)

def table(d, rows, widths=None, header=False):
    from docx.shared import Cm
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = 'Table Grid'
    for i, row in enumerate(rows):
        for j, segs in enumerate(row):
            c = t.cell(i, j)
            c.paragraphs[0].text = ''
            if isinstance(segs, str):
                segs = [N(segs)]
            for text, kind in segs:
                r = c.paragraphs[0].add_run(text)
                r.font.size = Pt(9)
                if kind == 'f':
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                if kind == 'b' or (header and i == 0) or (j == 0 and len(rows[0]) == 2):
                    r.bold = True
            if widths:
                c.width = Cm(widths[j])
    d.add_paragraph()
    return t

def footer(d, text):
    fp = d.sections[0].footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(text); r.font.size = Pt(8); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

def build_sepa():
    d = new_doc()
    intestazione(d)
    para(d, [N('MANDATO PER ADDEBITO DIRETTO SEPA CORE (SDD)')], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12)
    para(d, [N('Reg. UE 260/2012 – SEPA Core Direct Debit Scheme Rulebook (EPC)')], align=WD_ALIGN_PARAGRAPH.CENTER, size=8, space_after=10)
    table(d, [['Riferimento del mandato', [F('RIF_MANDATO_SEPA')]],
              ['Tipo di pagamento', 'Ricorrente']], widths=[6, 10.6])
    para(d, [N('Creditore')], bold=True)
    table(d, [['Denominazione', 'Studio Commercialista Maurizio Magotti'],
              ['Codice identificativo del creditore (Creditor Identifier)', [F('SEPA_CID')]],
              ['Indirizzo', 'Corso Vittorio Emanuele II, 101 – 46100 Mantova (MN) – Italia']], widths=[6, 10.6])
    para(d, [N('Debitore')], bold=True)
    table(d, [['Intestatario del conto', [F('INTESTATARIO_CONTO')]],
              ['Codice fiscale / P.IVA intestatario', [F('CF_INTESTATARIO')]],
              ['Indirizzo', [F('INDIRIZZO_COMPLETO')]],
              ['IBAN del conto di addebito', [F('IBAN_CLIENTE')]],
              ['BIC (facoltativo)', [F('BIC_CLIENTE')]],
              ['Sottoscrittore del mandato (se diverso dall\'intestatario)', [F('SOTTOSCRITTORE_SEPA')]]], widths=[6, 10.6])
    para(d, [N("Con la sottoscrizione del presente mandato il debitore autorizza (A) il creditore a richiedere alla banca del debitore l'addebito del proprio conto e (B) la banca del debitore a procedere a tale addebito conformemente alle disposizioni impartite dal creditore. Il debitore ha diritto di essere rimborsato dalla propria banca secondo gli accordi e alle condizioni che regolano il rapporto con quest'ultima. Se del caso, il rimborso deve essere richiesto nel termine di 8 settimane a decorrere dalla data di addebito in conto.")]).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para(d, [N("Il creditore comunica al debitore l'importo e la data di ciascun addebito con la relativa fattura (pre-notification), almeno 5 giorni di calendario prima della scadenza, salvo diverso accordo. I diritti del debitore sono indicati nella documentazione ottenibile dalla propria banca.")]).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para(d, [N('Rapporto sottostante: lettera di incarico professionale del '), F('DATA_LETTERA'), N(' (art. 7).')], space_after=14)
    table(d, [['Luogo e data', 'Firma del debitore / sottoscrittore'],
              ['\n\n____________________', '\n\n________________________']], widths=[8.3, 8.3], header=True)
    footer(d, 'Mandato SEPA Core – {{DENOMINAZIONE}}')
    d.save(OUT + 'MOD_mandato_SEPA_campi.docx')

def build_privacy():
    d = new_doc()
    intestazione(d)
    para(d, [N('INFORMATIVA SUL TRATTAMENTO DEI DATI PERSONALI DEI CLIENTI')], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12)
    para(d, [N('art. 13 Reg. UE 2016/679 (GDPR) e D.Lgs. 30 giugno 2003, n. 196')], align=WD_ALIGN_PARAGRAPH.CENTER, size=8, space_after=10)
    para(d, [N('Spett.le '), F('DENOMINAZIONE')])
    def sez(t, testo):
        para(d, [N(t)], bold=True, space_after=2)
        if isinstance(testo, str):
            testo = [N(testo)]
        para(d, testo).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sez('1. Titolare del trattamento', "Maurizio Magotti, Studio Commercialista Maurizio Magotti, Corso Vittorio Emanuele II, 101 – 46100 Mantova; tel. 0376 360750; e-mail info@studiomagotti.com; PEC maurizio.magotti@pec.it.")
    sez('2. Responsabile della protezione dei dati (DPO)', [F('TXT_DPO')])
    sez('3. Dati trattati', "Dati anagrafici e di contatto del Cliente, dei legali rappresentanti, dei soci e dei titolari effettivi; dati contabili, fiscali, patrimoniali e bancari; dati relativi a dipendenti, fornitori e clienti del Cliente contenuti nella documentazione consegnata; dati relativi a condanne penali e reati (art. 10 GDPR) nei soli limiti richiesti dalla normativa antiriciclaggio. Categorie particolari di dati (art. 9 GDPR) sono trattate solo se strettamente necessarie alla prestazione (es. oneri deducibili, agevolazioni).")
    table(d, [['Finalità', 'Base giuridica', 'Conservazione'],
              ['Esecuzione dell\'incarico professionale e attività precontrattuali', 'art. 6, par. 1, lett. b), GDPR', '10 anni dalla cessazione dell\'incarico (art. 2946 c.c.; art. 2220 c.c.)'],
              ['Adempimenti di legge: fiscali, contabili, societari, antiriciclaggio (D.Lgs. 231/2007), deontologici', 'art. 6, par. 1, lett. c), GDPR; art. 10 GDPR e art. 2-octies D.Lgs. 196/2003', '10 anni (artt. 31-32 D.Lgs. 231/2007; art. 2220 c.c.); termini di accertamento (art. 43 DPR 600/1973)'],
              ['Tutela dei diritti in sede giudiziale, recupero crediti', 'art. 6, par. 1, lett. f), GDPR (legittimo interesse)', 'Fino al passaggio in giudicato e ai termini di prescrizione'],
              ['Fatturazione e gestione amministrativa del rapporto', 'art. 6, par. 1, lett. b) e c), GDPR', '10 anni (art. 2220 c.c.)']], widths=[6, 5.3, 5.3], header=True)
    sez('4. Natura del conferimento', "Il conferimento dei dati è necessario per l'esecuzione dell'incarico e per l'adempimento degli obblighi di legge. Il rifiuto comporta l'impossibilità di instaurare o proseguire il rapporto; in ambito antiriciclaggio comporta l'obbligo di astensione (art. 42 D.Lgs. 231/2007).")
    sez('5. Destinatari', "Collaboratori e dipendenti autorizzati al trattamento (art. 29 GDPR; art. 2-quaterdecies D.Lgs. 196/2003); professionisti esterni coinvolti nell'incarico (notaio, avvocato, consulente del lavoro) previo accordo con il Cliente; fornitori di servizi informatici, gestionali, di firma digitale e conservazione e di archiviazione cloud, nominati responsabili del trattamento (art. 28 GDPR); Amministrazione finanziaria, Registro Imprese, enti previdenziali, autorità giudiziarie e di vigilanza, UIF, Ordine professionale, nei casi previsti dalla legge; istituti bancari per la gestione dei pagamenti. I dati non sono diffusi.")
    sez('6. Trasferimento extra UE', [N("Alcuni servizi cloud utilizzati dallo Studio possono comportare il trasferimento di dati verso Paesi terzi (USA). Il trasferimento avviene sulla base della decisione di adeguatezza della Commissione europea del 10 luglio 2023 (EU-U.S. Data Privacy Framework) per i fornitori certificati, ovvero delle clausole contrattuali standard (art. 46, par. 2, lett. c), GDPR). "), F('TXT_EXTRA_UE')])
    sez('7. Diritti dell\'interessato', "L'interessato può esercitare i diritti di accesso, rettifica, cancellazione, limitazione, portabilità e opposizione (artt. 15-22 GDPR) scrivendo ai recapiti del Titolare, nei limiti previsti dall'art. 2-undecies D.Lgs. 196/2003 e dagli obblighi di conservazione di legge. L'interessato ha diritto di proporre reclamo al Garante per la protezione dei dati personali (art. 77 GDPR; www.garanteprivacy.it).")
    sez('8. Processi decisionali automatizzati', "Non sono adottati processi decisionali automatizzati né profilazione ai sensi dell'art. 22 GDPR. La profilatura del rischio antiriciclaggio è svolta dal Professionista e non produce effetti esclusivamente automatizzati.")
    para(d, [N('Mantova, '), F('DATA_LETTERA')], space_after=10)
    para(d, [N('Il Titolare del trattamento – Maurizio Magotti - Commercialista e Revisore Legale')])
    para(d, [N('Per presa visione – Il Cliente  ________________________')], space_after=0)
    footer(d, 'Informativa privacy clienti – {{DENOMINAZIONE}}')
    d.save(OUT + 'MOD_informativa_privacy_campi.docx')

if __name__ == '__main__':
    build_mandato()
    build_aml()
    build_sepa()
    build_privacy()
    print('ok')
