"""Genera MOD_invio_mandati.xlsx (da salvare come .xlsm dopo l'import del modulo VBA)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import datetime

OUT = 'modelli/incarichi/automazione/MOD_invio_mandati.xlsx'
SINO = ['SI', 'NO']
PER = ['Mensile', 'Trimestrale', 'Semestrale', 'Annuale', 'Una tantum', 'A consuntivo']
MOD = ['Forfait', 'A tempo', 'Parametri D.M. 140/2012', 'Compreso nel forfait']

# (chiave, gruppo, descrizione, lista, larghezza, esempio, testo?)
C = []
def add(k, g, d, lst=None, w=16, ex='', txt=False):
    C.append((k, g, d, lst, w, ex, txt))

g = 'CONTROLLO'
add('STATO', g, 'DA INVIARE = elabora; ESEMPIO/SOSPESO = ignora. Aggiornato dalla macro', ['DA INVIARE', 'INVIATO', 'BOZZA CREATA', 'IN ANTEPRIMA', 'ERRORE', 'SOSPESO', 'ESEMPIO'], 14, 'ESEMPIO')
add('DATA_INVIO', g, 'Compilato dalla macro', None, 16)
add('ESITO', g, 'Compilato dalla macro (errori di validazione o esito invio)', None, 40)
add('CARTELLA_DOCUMENTI', g, 'Compilato dalla macro', None, 30)
g = 'ANAGRAFICA'
add('CODICE_CLIENTE', g, 'Codice anonimo (es. CL001): nome cartella e file', None, 12, 'CL000', True)
add('DENOMINAZIONE', g, 'Ragione sociale o nome e cognome', None, 28, 'ALFA S.R.L. (ESEMPIO)')
add('TIPO_CLIENTE', g, 'Determina legale rappresentante, foro e interessi moratori', ['Società di capitali', 'Società di persone', 'Impresa individuale', 'Professionista', 'Consumatore'], 20, 'Società di capitali')
add('INDIRIZZO', g, 'Via e numero civico', None, 26, 'Via Roma, 1')
add('CAP', g, '', None, 7, '46100', True)
add('CITTA', g, '', None, 14, 'Mantova')
add('PROV', g, 'Sigla', None, 6, 'MN')
add('CF', g, 'Codice fiscale', None, 18, '00000000000', True)
add('PIVA', g, 'Vuoto = "—"', None, 14, '00000000000', True)
add('PEC_CLIENTE', g, 'Se vuota, le comunicazioni art. 13 vanno all\'e-mail', None, 24, 'alfa@pec.esempio.it')
add('LEGALE_RAPP', g, 'Obbligatorio per le società', None, 20, 'Mario Rossi')
g = 'INVIO E-MAIL'
add('EMAIL_DESTINATARIO', g, 'Più indirizzi separati da ;', None, 26, 'amministrazione@esempio.it')
add('EMAIL_CC', g, 'Facoltativo', None, 20)
add('SALUTO', g, 'Apertura della cover. Vuoto = "Spett.le DENOMINAZIONE"', None, 22, 'Gentile Sig. Rossi')
g = 'CONDIZIONI DEL MANDATO'
add('DATA_LETTERA', g, 'Vuoto = data odierna', None, 12, datetime.date(2026, 11, 7))
add('DURATA', g, 'Art. 9', ['annuale', 'fino al completamento della prestazione'], 18, 'annuale')
add('GG_DISDETTA', g, 'Art. 9 – giorni di preavviso', None, 8, 60)
add('GG_CONSEGNA_DOC', g, 'Art. 4 – giorno del mese successivo', None, 8, 15)
add('DATA_CONSEGNA_DICH', g, 'Art. 4 – formato GG/MM', None, 9, '31/03', True)
add('COMPLESSITA', g, 'Art. 5', ['basso', 'medio', 'elevato'], 9, 'medio')
add('MOTIVAZIONE_COMPLESSITA', g, 'Art. 5 – motivazione sintetica', None, 28, 'volume delle registrazioni e numero di adempimenti periodici')
add('FATTURAZIONE', g, 'Art. 7', ['mensile', 'trimestrale', 'semestrale', 'annuale', 'a conclusione della prestazione'], 14, 'trimestrale')
add('GG_PAGAMENTO', g, 'Art. 7 – giorni da ricevimento fattura', None, 8, 30)
add('N_FATTURE_RECESSO', g, 'Art. 7 – fatture insolute per recesso', ['due', 'tre'], 8, 'due')
add('AGGIORNAMENTO_COMPENSI', g, 'Art. 5', ["in misura pari alla variazione dell'indice ISTAT FOI", 'secondo accordo tra le parti'], 26, "in misura pari alla variazione dell'indice ISTAT FOI")
add('DATA_COMUNICAZIONE_AGG', g, 'Art. 5 – formato GG/MM', None, 9, '30/11', True)
add('SPESE_GENERALI_PERC', g, 'Art. 5 – % rimborso forfetario; vuoto = non previsto', None, 9, '')
add('ACCONTO_IMPORTO', g, 'Art. 6 – importo; vuoto = vedi ACCONTO_PERC', None, 10, '')
add('ACCONTO_PERC', g, 'Art. 6 – %; entrambi vuoti = nessun acconto', None, 9, '')
g = 'CONTROLLI'
add('COMMITTENTE_FORTE', g, 'Banca, assicurazione, impresa >50 dip. o ricavi >10 mln (L. 49/2023)', SINO, 10, 'NO')
add('VERIFICA_EQUO_COMPENSO', g, 'SI = compenso verificato ≥ parametri D.M. 140/2012 (obbligatorio se committente forte)', SINO, 10, 'NO')
add('CLIENTE_REVISIONE', g, 'SI = cliente con incarico di revisione/sindaco-revisore dello Studio (art. 10 D.Lgs. 39/2010)', SINO, 10, 'NO')
add('CLAUSOLA_ART28', g, 'SI = inserisce clausola responsabile del trattamento (art. 28 GDPR)', SINO, 10, 'NO')

def prest(code, label, extra, ex):
    gg = 'PRESTAZIONE: ' + label
    add('P_' + code, gg, 'SI = conferita', SINO, 8, ex[0])
    for e in extra:
        add(*((e[0], gg) + e[1:]))
    add('MOD_' + code, gg, 'Modalità di compenso', MOD, 14, ex[1])
    add('IMP_' + code, gg, 'Importo per periodo (€, oltre cassa e IVA) o testo', None, 10, ex[2])
    add('PER_' + code, gg, 'Mensile/Trim./Sem./Annuale concorrono al totale annuo', PER, 12, ex[3])

prest('CONT', 'CONTABILITÀ', [('REGIME_CONTABILE', 'Regime', ['ordinaria', 'semplificata'], 12, 'ordinaria')], ['SI', 'Forfait', 250, 'Mensile'])
prest('BIL', 'BILANCIO', [('SCHEMA_BILANCIO', 'Schema', ['ordinaria', 'abbreviata (art. 2435-bis c.c.)', 'semplificata per le micro-imprese (art. 2435-ter c.c.)'], 18, 'abbreviata (art. 2435-bis c.c.)'),
                          ('RELAZIONE_GESTIONE', 'SI = include relazione sulla gestione', SINO, 9, 'NO')], ['SI', 'Forfait', 1500, 'Annuale'])
prest('DICH', 'DICHIARAZIONI', [('DICH_770', 'SI = 770 autonomi/capitale', SINO, 9, 'NO')], ['SI', 'Forfait', 1200, 'Annuale'])
prest('F24', 'IMPOSTE / F24 / ISA', [('CPB', 'SI = adempimenti CPB', SINO, 8, 'NO')], ['NO', '', '', ''])
prest('SOC', 'SOCIETARIO', [], ['SI', 'Forfait', 400, 'Annuale'])
prest('CONS', 'CONSULENZA', [], ['NO', '', '', ''])
prest('STRA', 'OP. STRAORDINARIA', [('DESCR_STRA', 'Tipo operazione', None, 20, '')], ['NO', '', '', ''])
prest('PER', 'PERIZIA / CTP', [('DESCR_PER', 'Oggetto', None, 20, '')], ['NO', '', '', ''])
prest('ALT', 'ALTRO', [('DESCR_ALT', 'Descrizione', None, 20, '')], ['NO', '', '', ''])
g = 'PAGAMENTO'
add('PAGAMENTO', g, 'Bonifico (IBAN Studio) o SDD (genera mandato SEPA)', ['Bonifico', 'SDD'], 10, 'Bonifico')
add('IBAN_CLIENTE', g, 'Obbligatorio se SDD', None, 30, '', True)
add('BIC_CLIENTE', g, 'Facoltativo', None, 12, '', True)
add('INTESTATARIO_CONTO', g, 'Vuoto = DENOMINAZIONE', None, 20)
add('CF_INTESTATARIO', g, 'Vuoto = P.IVA o C.F. cliente', None, 16, '', True)
g = 'ALLEGATI'
add('ALL_AML', g, 'SI = allega M1 + M3 antiriciclaggio', SINO, 8, 'SI')
add('ALL_PRIVACY', g, 'SI = allega informativa privacy', SINO, 8, 'SI')

GROUP_COLORS = ['1F3864', '2E75B6', '548235', 'BF8F00', 'C00000', '7030A0', '305496', '375623', '806000', '833C0B']
thin = Side(style='thin', color='BFBFBF')
B = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
# ---------------- Istruzioni
wi = wb.active; wi.title = 'Istruzioni'
istr = [
 ('AUTOMAZIONE INVIO LETTERE DI INCARICO – Studio Commercialista Maurizio Magotti', 'h'),
 ('', ''),
 ('1. INSTALLAZIONE (una tantum)', 'b'),
 ('a) Copiare nella stessa cartella locale (non OneDrive web): questo file, i 4 modelli MOD_*_campi.docx e modInvioMandati.bas.', ''),
 ('b) Aprire questo file in Excel → ALT+F11 (editor VBA) → File → Importa file… → modInvioMandati.bas → chiudere l\'editor.', ''),
 ('c) File → Salva con nome → tipo "Cartella di lavoro con attivazione macro di Excel (*.xlsm)" → es. Invio_mandati.xlsm.', ''),
 ('d) Facoltativo: scheda Sviluppo → Inserisci → Pulsante (controllo modulo) nel foglio Clienti, assegnando le macro InviaMandati, InviaRigaSelezionata, GeneraSoloDocumenti, VerificaDati.', ''),
 ('e) Outlook desktop deve avere configurato l\'account indicato in Impostazioni → ACCOUNT_MITTENTE.', ''),
 ('f) Compilare TUTTE le voci [DC] del foglio Impostazioni (polizza RC, cassa, tariffe, IBAN, SEPA_CID): finché restano [DC] la macro blocca l\'invio.', ''),
 ('', ''),
 ('2. UTILIZZO', 'b'),
 ('a) Foglio Clienti: una riga per cliente dalla riga 4. La riga 2 contiene i nomi tecnici dei campi: NON modificarli. La riga 3 contiene le istruzioni di compilazione.', ''),
 ('b) Impostare STATO = "DA INVIARE" sulle righe da elaborare (la riga di esempio CL000 ha STATO = ESEMPIO e viene ignorata).', ''),
 ('c) Eseguire (ALT+F8): VerificaDati → controlla senza generare; GeneraSoloDocumenti → crea DOCX/PDF senza inviare; InviaMandati → genera e invia; InviaRigaSelezionata → solo la riga attiva.', ''),
 ('d) Primo utilizzo: impostare MODALITA_INVIO = ANTEPRIMA o BOZZA e controllare PDF ed e-mail; poi passare a INVIO.', ''),
 ('e) Output: CARTELLA_OUTPUT\\<CODICE_CLIENTE>\\AAAA-MM-GG_<CODICE>_lettera_incarico.pdf (+ antiriciclaggio, mandato_SEPA, informativa_privacy). PDF/A se PDF_A = SI.', ''),
 ('f) Dopo l\'invio: STATO = INVIATO, DATA_INVIO e ESITO compilati, riga nel foglio Log. Per reinviare reimpostare STATO = DA INVIARE.', ''),
 ('', ''),
 ('3. CONTROLLI BLOCCANTI (la riga va in ERRORE e non viene inviata)', 'b'),
 ('• Campi obbligatori mancanti; e-mail o IBAN formalmente non validi; nessuna prestazione conferita.', ''),
 ('• Prestazione conferita senza modalità/importo/periodicità (preventivo obbligatorio: art. 9, c. 4, D.L. 1/2012).', ''),
 ('• Polizza RC, cassa e tariffe non compilate in Impostazioni (art. 9, c. 4, D.L. 1/2012; art. 5 DPR 137/2012).', ''),
 ('• CLIENTE_REVISIONE = SI con tenuta contabilità o bilancio conferiti (indipendenza: art. 10 D.Lgs. 39/2010).', ''),
 ('• COMMITTENTE_FORTE = SI senza VERIFICA_EQUO_COMPENSO = SI (L. 49/2023; parametri D.M. 140/2012).', ''),
 ('• Pagamento SDD senza IBAN cliente o SEPA_CID dello Studio.', ''),
 ('• Campi {{…}} non compilati o voci [DC] residue nei documenti generati.', ''),
 ('', ''),
 ('4. LOGICA DEI MODELLI WORD', 'b'),
 ('• {{CAMPO}} = sostituito con la colonna omonima del foglio Clienti o la voce omonima di Impostazioni (o con un campo calcolato dalla macro).', ''),
 ('• {{?FLAG}} a inizio paragrafo / riga di tabella = il paragrafo o la riga sono mantenuti solo se la condizione è vera (es. P_CONT = SI).', ''),
 ('• Per aggiungere un campo: inserire {{NUOVO_CAMPO}} nel modello Word (in un\'unica digitazione, senza formattazioni intermedie) e aggiungere la colonna NUOVO_CAMPO in riga 2 del foglio Clienti.', ''),
 ('• Campi calcolati: INDIRIZZO_COMPLETO, FRASE_LEGALE_RAPP, TXT_FIRMA_LR, TXT_RELAZIONE_GESTIONE, TXT_770, TXT_CPB, TXT_SPESE_GENERALI, TXT_ACCONTO, TXT_MODALITA_PAGAMENTO, TXT_FORO, TOTALE_ANNUO, RECAPITO_COMUNICAZIONI, TERMINE_RESTITUZIONE, RIF_MANDATO_SEPA, FIRMATARIO, QUALITA_DICHIARANTE.', ''),
 ('• TOTALE_ANNUO = somma IMP_ delle prestazioni conferite × periodi (Mensile 12, Trimestrale 4, Semestrale 2, Annuale 1); Una tantum e A consuntivo esclusi.', ''),
 ('', ''),
 ('5. AVVERTENZE', 'b'),
 ('⚠️ L\'invio con e-mail ordinaria non dà prova legale di consegna: per contestazioni sul compenso conservare la lettera controfirmata restituita dal cliente (art. 2233 c.c.; art. 9, c. 4, D.L. 1/2012).', ''),
 ('⚠️ L\'adeguata verifica antiriciclaggio deve essere completata prima dell\'avvio della prestazione (art. 42 D.Lgs. 231/2007): l\'invio della dichiarazione M3 non sostituisce identificazione e profilatura (MOD_AML_valutazione_rischio.xlsx).', ''),
 ('⚠️ Riservatezza: questo file contiene dati di clienti. Non salvarlo nel repository; conservarlo nell\'archivio dello Studio.', ''),
]
for i, (t, k) in enumerate(istr, start=1):
    c = wi.cell(row=i, column=1, value=t)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    if k == 'h': c.font = Font(bold=True, size=13, color='1F3864')
    elif k == 'b': c.font = Font(bold=True, color='1F3864')
wi.column_dimensions['A'].width = 140

# ---------------- Clienti
ws = wb.create_sheet('Clienti')
groups = []
for idx, (k, g_, d, lst, w, ex, txt) in enumerate(C, start=1):
    if not groups or groups[-1][0] != g_:
        groups.append([g_, idx, idx])
    else:
        groups[-1][2] = idx
for gi, (g_, a, b) in enumerate(groups):
    col = GROUP_COLORS[gi % len(GROUP_COLORS)] if not g_.startswith('PRESTAZIONE') else ('375623' if gi % 2 else '548235')
    if b > a: ws.merge_cells(start_row=1, start_column=a, end_row=1, end_column=b)
    c = ws.cell(row=1, column=a, value=g_)
    c.font = Font(bold=True, color='FFFFFF'); c.alignment = Alignment(horizontal='center', vertical='center')
    for j in range(a, b + 1):
        ws.cell(row=1, column=j).fill = PatternFill('solid', fgColor=col)
        ws.cell(row=2, column=j).fill = PatternFill('solid', fgColor='D9E1F2')
        ws.cell(row=3, column=j).fill = PatternFill('solid', fgColor='F2F2F2')
dvs = {}
for idx, (k, g_, d, lst, w, ex, txt) in enumerate(C, start=1):
    L = get_column_letter(idx)
    h = ws.cell(row=2, column=idx, value=k); h.font = Font(bold=True, size=9); h.border = B
    h.alignment = Alignment(wrap_text=True, vertical='center')
    dd = ws.cell(row=3, column=idx, value=d); dd.font = Font(italic=True, size=8, color='595959'); dd.border = B
    dd.alignment = Alignment(wrap_text=True, vertical='top')
    ws.column_dimensions[L].width = w
    e = ws.cell(row=4, column=idx, value=ex if ex != '' else None)
    e.font = Font(color='7F7F7F', italic=True); e.border = B
    if txt:
        for r in range(4, 504): ws.cell(row=r, column=idx).number_format = '@'
    if isinstance(ex, datetime.date) or k == 'DATA_INVIO':
        for r in range(4, 504): ws.cell(row=r, column=idx).number_format = 'dd/mm/yyyy' if k != 'DATA_INVIO' else 'dd/mm/yyyy hh:mm'
    if k.startswith('IMP_') or k == 'ACCONTO_IMPORTO':
        for r in range(4, 504): ws.cell(row=r, column=idx).number_format = '#,##0.00'
    if lst:
        key = tuple(lst)
        if key not in dvs:
            # lista su foglio Liste per evitare il limite di 255 caratteri
            dvs[key] = None
        dvs.setdefault(key, None)
ws.row_dimensions[2].height = 30
ws.row_dimensions[3].height = 60
ws.freeze_panes = 'G4'

# Liste per convalida
wl = wb.create_sheet('Liste')
listcols = {}
for idx, (k, g_, d, lst, w, ex, txt) in enumerate(C, start=1):
    if lst and tuple(lst) not in listcols:
        n = len(listcols) + 1
        for i, v in enumerate(lst, start=1):
            wl.cell(row=i, column=n, value=v)
        listcols[tuple(lst)] = (n, len(lst))
for idx, (k, g_, d, lst, w, ex, txt) in enumerate(C, start=1):
    if lst:
        n, ln = listcols[tuple(lst)]
        L = get_column_letter(n)
        dv = DataValidation(type='list', formula1=f"=Liste!${L}$1:${L}${ln}", allow_blank=True,
                            showErrorMessage=k not in ('STATO',), errorTitle='Valore non ammesso',
                            error='Selezionare un valore dall\'elenco.')
        dv.add(f'{get_column_letter(idx)}4:{get_column_letter(idx)}503')
        ws.add_data_validation(dv)
wl.sheet_state = 'hidden'

# ---------------- Impostazioni
wsI = wb.create_sheet('Impostazioni')
imp = [
 ('CHIAVE', 'VALORE', 'NOTE'),
 ('MODALITA_INVIO', 'ANTEPRIMA', 'INVIO = invio automatico; BOZZA = salva in Bozze di Outlook; ANTEPRIMA = apre la mail senza inviare. Primo utilizzo: ANTEPRIMA'),
 ('CONFERMA_PRIMA_INVIO', 'SI', 'SI = richiesta di conferma prima dell\'invio massivo'),
 ('ACCOUNT_MITTENTE', 'info@studiomagotti.com', 'Indirizzo SMTP dell\'account Outlook da cui inviare'),
 ('CCN_ARCHIVIO', '', 'Facoltativo: indirizzo in copia nascosta per archivio'),
 ('GG_TERMINE_RESTITUZIONE', 15, 'Giorni dalla DATA_LETTERA per la restituzione firmata (citati nella cover)'),
 ('CARTELLA_MODELLI', '', 'Vuoto = cartella di questo file. Obbligatorio se il file è su OneDrive/SharePoint'),
 ('CARTELLA_OUTPUT', '', 'Vuoto = <cartella file>\\Mandati_generati. Consigliato: archivio master dello Studio'),
 ('MODELLO_MANDATO', 'MOD_mandato_campi.docx', 'Nome file del modello lettera di incarico'),
 ('MODELLO_AML', 'MOD_AML_cliente_campi.docx', 'Informativa M1 + dichiarazione M3'),
 ('MODELLO_SEPA', 'MOD_mandato_SEPA_campi.docx', 'Mandato SDD (usato solo se PAGAMENTO = SDD)'),
 ('MODELLO_PRIVACY', 'MOD_informativa_privacy_campi.docx', 'Informativa art. 13 GDPR'),
 ('SALVA_DOCX', 'SI', 'SI = salva anche la versione Word compilata'),
 ('PDF_A', 'SI', 'SI = esporta in PDF/A (ISO 19005-1)'),
 ('CASSA_PREVIDENZA', '[DC]', 'CNPR (ragionieri) o CNPADC (dottori commercialisti) – contributo integrativo 4%'),
 ('TARIFFA_TITOLARE', '[DC]', 'Tariffa oraria titolare (€) – art. 5'),
 ('TARIFFA_COLLABORATORI', '[DC]', 'Tariffa oraria collaboratori (€) – art. 5'),
 ('IBAN_STUDIO', '[DC]', 'IBAN per bonifico – art. 7'),
 ('SEPA_CID', '[DC]', 'Creditor Identifier SEPA rilasciato dalla banca dello Studio (es. IT00ZZZ…)'),
 ('RC_COMPAGNIA', '[DC]', 'Polizza RC – compagnia (art. 9, c. 4, D.L. 1/2012; art. 5 DPR 137/2012)'),
 ('RC_POLIZZA', '[DC]', 'Polizza RC – numero'),
 ('RC_MASS_SINISTRO', '[DC]', 'Massimale per sinistro (€)'),
 ('RC_MASS_ANNUO', '[DC]', 'Massimale aggregato annuo (€)'),
 ('RC_SCADENZA', '[DC]', 'Scadenza polizza (GG/MM/AAAA)'),
 ('RC_SINDACO_REVISORE', '[DC]', 'Copertura incarichi sindaco/revisore: es. "Inclusa" / "Non inclusa"'),
 ('TXT_DPO', "Il Titolare non ha designato un Responsabile della protezione dei dati, non ricorrendone i presupposti di cui all'art. 37 Reg. UE 2016/679.", '⚠️ Verificare (CLAUDE.md: DPO [DC])'),
 ('TXT_EXTRA_UE', 'Le informazioni sui fornitori e sulle garanzie adottate sono disponibili su richiesta al Titolare.', 'Informativa privacy, sez. 6'),
]
for r, row in enumerate(imp, start=1):
    for c, v in enumerate(row, start=1):
        cell = wsI.cell(row=r, column=c, value=v)
        cell.border = B
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        if r == 1:
            cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill('solid', fgColor='1F3864')
        elif v == '[DC]':
            cell.fill = PatternFill('solid', fgColor='FFF2CC')
wsI.column_dimensions['A'].width = 26; wsI.column_dimensions['B'].width = 50; wsI.column_dimensions['C'].width = 80
dvm = DataValidation(type='list', formula1='"INVIO,BOZZA,ANTEPRIMA"'); dvm.add('B2'); wsI.add_data_validation(dvm)
dvs2 = DataValidation(type='list', formula1='"SI,NO"'); dvs2.add('B3'); dvs2.add('B13'); dvs2.add('B14'); wsI.add_data_validation(dvs2)
wsI.freeze_panes = 'A2'

# ---------------- Cover
wc = wb.create_sheet('Cover')
corpo = """{{SALUTO}},

facendo seguito ai colloqui intercorsi, trasmettiamo in allegato la lettera di incarico professionale e il preventivo di massima, redatti ai sensi dell'art. 9, c. 4, D.L. 24 gennaio 2012, n. 1, relativi alle prestazioni che lo Studio svolgerà in favore di {{DENOMINAZIONE}}.

Documenti allegati:
{{ELENCO_ALLEGATI}}
Vi preghiamo di restituirci entro il {{TERMINE_RESTITUZIONE}}:
{{ELENCO_RESTITUZIONE}}
La documentazione può essere restituita firmata digitalmente all'indirizzo PEC maurizio.magotti@pec.it, oppure sottoscritta con firma autografa e trasmessa in scansione all'indirizzo info@studiomagotti.com, con successiva consegna degli originali allo Studio.

Ai sensi dell'art. 42 D.Lgs. 21 novembre 2007, n. 231, lo svolgimento della prestazione potrà avere inizio solo al completamento dell'adeguata verifica antiriciclaggio.

Restiamo a disposizione per ogni chiarimento.

Distinti saluti.

Maurizio Magotti - Commercialista e Revisore Legale
Studio Commercialista Maurizio Magotti
Corso Vittorio Emanuele II, 101 – 46100 Mantova
Tel. +39 0376 360750 – info@studiomagotti.com – PEC maurizio.magotti@pec.it"""
disc = ("Il presente messaggio e i relativi allegati contengono informazioni riservate e coperte da segreto professionale, "
        "destinate esclusivamente ai destinatari indicati. Se ricevuto per errore, si prega di darne comunicazione al mittente "
        "e di cancellarlo; ogni utilizzo, copia o diffusione non autorizzati sono vietati (art. 616 c.p.; Reg. UE 2016/679).")
cov = [('CHIAVE', 'TESTO (campi {{...}} come nel foglio Clienti; {{ELENCO_ALLEGATI}} e {{ELENCO_RESTITUZIONE}} generati dalla macro)'),
       ('OGGETTO', 'Lettera di incarico professionale e preventivo – {{DENOMINAZIONE}}'),
       ('CORPO', corpo),
       ('DISCLAIMER', disc)]
for r, (a, b) in enumerate(cov, start=1):
    ca = wc.cell(row=r, column=1, value=a); cb = wc.cell(row=r, column=2, value=b)
    for c in (ca, cb):
        c.alignment = Alignment(wrap_text=True, vertical='top'); c.border = B
    if r == 1:
        for c in (ca, cb): c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
    else:
        ca.font = Font(bold=True)
wc.column_dimensions['A'].width = 14; wc.column_dimensions['B'].width = 120
wc.row_dimensions[3].height = 520; wc.row_dimensions[4].height = 45

# ---------------- Log
wlg = wb.create_sheet('Log')
for i, h in enumerate(['Data/ora', 'Utente', 'Codice', 'Denominazione', 'Destinatario', 'Fase', 'Esito', 'Dettaglio', 'File'], start=1):
    c = wlg.cell(row=1, column=i, value=h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
for L, w in zip('ABCDEFGHI', [17, 12, 10, 26, 26, 14, 9, 50, 60]):
    wlg.column_dimensions[L].width = w
for r in range(2, 2000): wlg.cell(row=r, column=1).number_format = 'dd/mm/yyyy hh:mm'
wlg.freeze_panes = 'A2'

wb.active = 1
wb.save(OUT)
print('ok', len(C), 'colonne')
