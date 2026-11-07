Attribute VB_Name = "modInvioMandati"
'==============================================================================
' AUTOMAZIONE INVIO LETTERE DI INCARICO
' Studio Commercialista Maurizio Magotti
'------------------------------------------------------------------------------
' Requisiti: Windows, Excel + Word + Outlook desktop (Office 2019 / Microsoft 365)
' Nessun riferimento da attivare (late binding).
' Fogli richiesti: Clienti, Impostazioni, Cover, Log.
' Macro pubbliche:
'   VerificaDati            controlla le righe "DA INVIARE" senza generare nulla
'   GeneraSoloDocumenti     genera DOCX/PDF delle righe "DA INVIARE" senza inviare
'   InviaMandati            genera i documenti e invia le e-mail (righe "DA INVIARE")
'   InviaRigaSelezionata    come InviaMandati, solo per la riga selezionata
'==============================================================================
Option Explicit

Private Const SH_CLIENTI As String = "Clienti"
Private Const SH_IMPOST As String = "Impostazioni"
Private Const SH_COVER As String = "Cover"
Private Const SH_LOG As String = "Log"
Private Const ROW_KEYS As Long = 2          ' riga con i nomi tecnici dei campi
Private Const ROW_FIRST As Long = 4         ' prima riga dati

' costanti Word / Outlook (late binding)
Private Const wdFindStop As Long = 0
Private Const wdCollapseEnd As Long = 0
Private Const wdWithInTable As Long = 12
Private Const wdExportFormatPDF As Long = 17
Private Const wdExportOptimizeForPrint As Long = 0
Private Const wdFormatXMLDocument As Long = 12
Private Const wdDoNotSaveChanges As Long = 0
Private Const wdNoHighlight As Long = 0
Private Const olMailItem As Long = 0

Private gSet As Object      ' impostazioni (Scripting.Dictionary)
Private gCols As Object     ' nome campo -> colonna
Private gWord As Object
Private gOutlook As Object

' sigle delle prestazioni: colonne P_x, MOD_x, IMP_x, PER_x
Private Const PRESTAZIONI As String = "CONT,BIL,DICH,F24,SOC,CONS,STRA,PER,ALT"

'==============================================================================
' MACRO PUBBLICHE
'==============================================================================
Public Sub VerificaDati()
    Esegui False, "VERIFICA"
End Sub

Public Sub GeneraSoloDocumenti()
    Esegui False, "DOCUMENTI"
End Sub

Public Sub InviaMandati()
    Esegui False, "INVIO"
End Sub

Public Sub InviaRigaSelezionata()
    Esegui True, "INVIO"
End Sub

'==============================================================================
' CICLO PRINCIPALE
'==============================================================================
Private Sub Esegui(ByVal soloSelezione As Boolean, ByVal fase As String)
    Dim ws As Worksheet, righe As Collection, v As Variant
    Dim lastRow As Long, lastCol As Long, r As Long
    Dim nOk As Long, nErr As Long, modalita As String, msg As String

    On Error GoTo Errore
    Set ws = ThisWorkbook.Worksheets(SH_CLIENTI)
    CaricaImpostazioni
    CaricaColonne ws
    lastCol = ws.Cells(ROW_KEYS, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, Col("CODICE_CLIENTE")).End(xlUp).Row

    Set righe = New Collection
    If soloSelezione Then
        If ActiveSheet.Name <> SH_CLIENTI Or ActiveCell.Row < ROW_FIRST Then
            MsgBox "Selezionare una cella nella riga del cliente (foglio Clienti).", vbExclamation
            Exit Sub
        End If
        righe.Add ActiveCell.Row
    Else
        For r = ROW_FIRST To lastRow
            If UCase$(Trim$(CStr(ws.Cells(r, Col("STATO")).Value))) = "DA INVIARE" Then righe.Add r
        Next r
    End If
    If righe.Count = 0 Then
        MsgBox "Nessuna riga con STATO = DA INVIARE.", vbInformation
        Exit Sub
    End If

    modalita = UCase$(Impostazione("MODALITA_INVIO"))
    If fase = "INVIO" Then
        If modalita <> "INVIO" And modalita <> "BOZZA" And modalita <> "ANTEPRIMA" Then
            MsgBox "Impostazioni: MODALITA_INVIO deve essere INVIO, BOZZA o ANTEPRIMA.", vbCritical
            Exit Sub
        End If
        If modalita = "INVIO" And UCase$(Impostazione("CONFERMA_PRIMA_INVIO")) <> "NO" Then
            If MsgBox("Saranno generate e INVIATE " & righe.Count & " e-mail dall'account " & _
                      Impostazione("ACCOUNT_MITTENTE") & "." & vbCrLf & _
                      "Le righe con errori di validazione saranno escluse." & vbCrLf & vbCrLf & _
                      "Procedere?", vbYesNo + vbQuestion, "Invio mandati") <> vbYes Then Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Elaborazione mandati..."
    For Each v In righe
        Application.StatusBar = "Elaborazione riga " & v & "..."
        ProcessaRiga ws, CLng(v), lastCol, fase, modalita, nOk, nErr
        DoEvents
    Next v

Fine:
    ChiudiWord
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If nOk + nErr > 0 Then
        msg = "Elaborazione completata." & vbCrLf & vbCrLf & _
              "Righe corrette: " & nOk & vbCrLf & "Righe in errore: " & nErr
        If nErr > 0 Then msg = msg & vbCrLf & vbCrLf & "Dettaglio nella colonna ESITO e nel foglio Log."
        MsgBox msg, IIf(nErr > 0, vbExclamation, vbInformation), "Invio mandati"
    End If
    Exit Sub

Errore:
    MsgBox "Errore: " & Err.Description, vbCritical, "Invio mandati"
    Resume Fine
End Sub

'==============================================================================
' ELABORAZIONE DI UNA RIGA
'==============================================================================
Private Sub ProcessaRiga(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, _
                         ByVal fase As String, ByVal modalita As String, _
                         ByRef nOk As Long, ByRef nErr As Long)
    Dim f As Object, flags As Object, errs As String, cartella As String, prefisso As String
    Dim allegati As Collection, descr As Collection, pdf As String, errDoc As String
    Dim esitoMail As String, statoFinale As String

    On Error GoTo ErroreRiga
    Set f = LeggiRiga(ws, r, lastCol)
    Set flags = CreateObject("Scripting.Dictionary")
    CalcolaDerivati f, flags
    errs = Valida(f, flags)
    If Len(errs) > 0 Then
        ScriviEsito ws, r, "ERRORE", errs, ""
        ScriviLog f, fase, "ERRORE", errs, ""
        nErr = nErr + 1
        Exit Sub
    End If

    If fase = "VERIFICA" Then
        ScriviEsito ws, r, "", "Verifica superata il " & Format$(Now, "dd/mm/yyyy hh:nn"), ""
        nOk = nOk + 1
        Exit Sub
    End If

    cartella = CartellaOutput() & "\" & NomeSicuro(f("CODICE_CLIENTE"))
    CreaCartella cartella
    prefisso = Format$(CDate(f("DATA_LETTERA")), "yyyy-mm-dd") & "_" & NomeSicuro(f("CODICE_CLIENTE")) & "_"
    Set allegati = New Collection
    Set descr = New Collection

    pdf = GeneraDocumento(PercorsoModello("MODELLO_MANDATO"), f, flags, cartella, prefisso & "lettera_incarico", errDoc)
    If Len(errDoc) > 0 Then GoTo ErroreDoc
    allegati.Add pdf: descr.Add "Lettera di incarico professionale e preventivo (art. 9, c. 4, D.L. 1/2012)"

    If SiNo(f("ALL_AML")) Then
        pdf = GeneraDocumento(PercorsoModello("MODELLO_AML"), f, flags, cartella, prefisso & "antiriciclaggio", errDoc)
        If Len(errDoc) > 0 Then GoTo ErroreDoc
        allegati.Add pdf: descr.Add "Informativa antiriciclaggio (M1) e dichiarazione del cliente (M3) - D.Lgs. 231/2007"
    End If
    If flags("SDD") Then
        pdf = GeneraDocumento(PercorsoModello("MODELLO_SEPA"), f, flags, cartella, prefisso & "mandato_SEPA", errDoc)
        If Len(errDoc) > 0 Then GoTo ErroreDoc
        allegati.Add pdf: descr.Add "Mandato per addebito diretto SEPA Core"
    End If
    If SiNo(f("ALL_PRIVACY")) Then
        pdf = GeneraDocumento(PercorsoModello("MODELLO_PRIVACY"), f, flags, cartella, prefisso & "informativa_privacy", errDoc)
        If Len(errDoc) > 0 Then GoTo ErroreDoc
        allegati.Add pdf: descr.Add "Informativa sul trattamento dei dati personali (art. 13 Reg. UE 2016/679)"
    End If

    If fase = "DOCUMENTI" Then
        ScriviEsito ws, r, "", "Documenti generati il " & Format$(Now, "dd/mm/yyyy hh:nn") & " (" & allegati.Count & " PDF)", cartella
        ScriviLog f, fase, "OK", "Documenti generati", ElencoFile(allegati)
        nOk = nOk + 1
        Exit Sub
    End If

    esitoMail = InviaEmail(f, allegati, descr, modalita)
    If Left$(esitoMail, 6) = "ERRORE" Then
        ScriviEsito ws, r, "ERRORE", esitoMail, cartella
        ScriviLog f, fase, "ERRORE", esitoMail, ElencoFile(allegati)
        nErr = nErr + 1
        Exit Sub
    End If
    Select Case modalita
        Case "INVIO": statoFinale = "INVIATO"
        Case "BOZZA": statoFinale = "BOZZA CREATA"
        Case Else: statoFinale = "IN ANTEPRIMA"
    End Select
    ScriviEsito ws, r, statoFinale, esitoMail, cartella
    If modalita = "INVIO" Then ws.Cells(r, Col("DATA_INVIO")).Value = Now
    ScriviLog f, fase & "/" & modalita, "OK", esitoMail, ElencoFile(allegati)
    nOk = nOk + 1
    Exit Sub

ErroreDoc:
    ScriviEsito ws, r, "ERRORE", errDoc, cartella
    ScriviLog f, fase, "ERRORE", errDoc, ""
    nErr = nErr + 1
    Exit Sub

ErroreRiga:
    ScriviEsito ws, r, "ERRORE", "Errore imprevisto: " & Err.Description, ""
    If Not f Is Nothing Then ScriviLog f, fase, "ERRORE", Err.Description, ""
    nErr = nErr + 1
End Sub

'==============================================================================
' LETTURA DATI
'==============================================================================
Private Function LeggiRiga(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As Object
    Dim f As Object, c As Long, k As String, v As Variant, s As String
    Set f = CreateObject("Scripting.Dictionary")
    f.CompareMode = vbTextCompare
    For c = 1 To lastCol
        k = UCase$(Trim$(CStr(ws.Cells(ROW_KEYS, c).Value)))
        If Len(k) > 0 Then
            v = ws.Cells(r, c).Value
            If IsError(v) Then
                s = ""
            ElseIf VarType(v) = vbDate Then
                If k = "DATA_CONSEGNA_DICH" Or k = "DATA_COMUNICAZIONE_AGG" Then
                    s = Format$(v, "dd/mm")
                Else
                    s = Format$(v, "dd/mm/yyyy")
                End If
            ElseIf IsNumeric(v) And Not IsEmpty(v) And (Left$(k, 4) = "IMP_" Or k = "ACCONTO_IMPORTO") Then
                s = Format$(CDbl(v), "#,##0.00")
            Else
                s = Trim$(CStr(v))
            End If
            f(k) = s
            f("RAW_" & k) = v
        End If
    Next c
    Set LeggiRiga = f
End Function

Private Sub CaricaImpostazioni()
    Dim ws As Worksheet, r As Long, k As String, v As Variant
    Set gSet = CreateObject("Scripting.Dictionary")
    gSet.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(SH_IMPOST)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        k = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Len(k) > 0 Then
            v = ws.Cells(r, 2).Value
            If VarType(v) = vbDate Then
                gSet(k) = Format$(v, "dd/mm/yyyy")
            ElseIf IsNumeric(v) And Not IsEmpty(v) And (Left$(k, 8) = "TARIFFA_" Or Left$(k, 7) = "RC_MASS") Then
                gSet(k) = Format$(CDbl(v), "#,##0.00")
            Else
                gSet(k) = Trim$(CStr(v))
            End If
        End If
    Next r
End Sub

Private Sub CaricaColonne(ByVal ws As Worksheet)
    Dim c As Long, k As String
    Set gCols = CreateObject("Scripting.Dictionary")
    gCols.CompareMode = vbTextCompare
    For c = 1 To ws.Cells(ROW_KEYS, ws.Columns.Count).End(xlToLeft).Column
        k = UCase$(Trim$(CStr(ws.Cells(ROW_KEYS, c).Value)))
        If Len(k) > 0 Then gCols(k) = c
    Next c
End Sub

Private Function Col(ByVal k As String) As Long
    If Not gCols.Exists(k) Then Err.Raise vbObjectError + 1, , "Colonna mancante nel foglio Clienti: " & k
    Col = gCols(k)
End Function

Private Function Impostazione(ByVal k As String) As String
    If gSet Is Nothing Then CaricaImpostazioni
    If gSet.Exists(k) Then Impostazione = gSet(k) Else Impostazione = ""
End Function

Private Function SiNo(ByVal v As String) As Boolean
    v = UCase$(Trim$(v))
    SiNo = (v = "SI" Or v = "S" & ChrW(204) Or v = "S" & ChrW(236) Or v = "X")
End Function

'==============================================================================
' CAMPI DERIVATI E CONDIZIONI
'==============================================================================
Private Sub CalcolaDerivati(ByVal f As Object, ByVal flags As Object)
    Dim tipo As String, isSoc As Boolean, isCons As Boolean, k As Variant, p As Variant
    Dim lr As String, totale As Double, molt As Double, gg As Long

    ' impostazioni dello Studio disponibili come campi
    For Each k In gSet.Keys
        If Not f.Exists(k) Then f(k) = gSet(k)
    Next k

    tipo = LCase$(f("TIPO_CLIENTE"))
    isSoc = (InStr(tipo, "societ") > 0)
    isCons = (InStr(tipo, "consumatore") > 0)
    lr = f("LEGALE_RAPP")

    If f("DATA_LETTERA") = "" Then f("DATA_LETTERA") = Format$(Date, "dd/mm/yyyy")
    f("INDIRIZZO_COMPLETO") = f("INDIRIZZO") & " " & ChrW(8211) & " " & f("CAP") & " " & f("CITTA") & _
                              IIf(f("PROV") <> "", " (" & f("PROV") & ")", "")
    If f("PIVA") = "" Then f("PIVA") = ChrW(8212)
    If f("PEC_CLIENTE") = "" Then
        f("RECAPITO_COMUNICAZIONI") = "e-mail " & f("EMAIL_DESTINATARIO")
        f("PEC_CLIENTE") = ChrW(8212)
    Else
        f("RECAPITO_COMUNICAZIONI") = "PEC " & f("PEC_CLIENTE")
    End If
    If f("SALUTO") = "" Then f("SALUTO") = "Spett.le " & f("DENOMINAZIONE")

    If isSoc And lr <> "" Then
        f("FRASE_LEGALE_RAPP") = ", in persona del legale rappresentante " & lr
        f("TXT_FIRMA_LR") = " " & ChrW(8211) & " il legale rappresentante " & lr
        f("FIRMATARIO") = lr
        f("QUALITA_DICHIARANTE") = "legale rappresentante di " & f("DENOMINAZIONE") & ", C.F. " & f("CF")
    Else
        f("FRASE_LEGALE_RAPP") = ""
        f("TXT_FIRMA_LR") = ""
        f("FIRMATARIO") = f("DENOMINAZIONE")
        f("QUALITA_DICHIARANTE") = "cliente in proprio"
    End If

    f("TXT_RELAZIONE_GESTIONE") = IIf(SiNo(f("RELAZIONE_GESTIONE")), " e della relazione sulla gestione", "")
    f("TXT_770") = IIf(SiNo(f("DICH_770")), ", 770 limitatamente ai redditi di lavoro autonomo e di capitale", "")
    f("TXT_CPB") = IIf(SiNo(f("CPB")), " e al concordato preventivo biennale", "")
    f("TXT_CPB_TAB") = IIf(SiNo(f("CPB")), ", CPB", "")

    If f("SPESE_GENERALI_PERC") <> "" Then
        f("TXT_SPESE_GENERALI") = ", e un rimborso forfetario delle spese generali pari al " & f("SPESE_GENERALI_PERC") & "% del compenso"
    Else
        f("TXT_SPESE_GENERALI") = ""
    End If

    If f("ACCONTO_IMPORTO") <> "" Then
        f("TXT_ACCONTO") = "All'accettazione dell'incarico il Cliente corrisponde un acconto di " & ChrW(8364) & " " & f("ACCONTO_IMPORTO") & "."
    ElseIf f("ACCONTO_PERC") <> "" Then
        f("TXT_ACCONTO") = "All'accettazione dell'incarico il Cliente corrisponde un acconto pari al " & f("ACCONTO_PERC") & "% del compenso preventivato."
    Else
        f("TXT_ACCONTO") = "Non " & ChrW(232) & " previsto il versamento di acconti."
    End If

    If UCase$(f("PAGAMENTO")) = "SDD" Then
        f("TXT_MODALITA_PAGAMENTO") = "addebito diretto SEPA Core (SDD) sul conto del Cliente IBAN " & f("IBAN_CLIENTE") & _
            ", sulla base del mandato sottoscritto dal Cliente (Allegato D); l'importo e la data di addebito sono preavvisati con la fattura"
    Else
        f("TXT_MODALITA_PAGAMENTO") = "bonifico bancario sul conto corrente intestato al Professionista, IBAN " & f("IBAN_STUDIO")
    End If

    If isCons Then
        f("TXT_FORO") = "Per ogni controversia " & ChrW(232) & " competente il foro di residenza o di domicilio elettivo del Cliente consumatore (art. 66-bis D.Lgs. 6 settembre 2005, n. 206)."
    Else
        f("TXT_FORO") = "Per ogni controversia " & ChrW(232) & " competente in via esclusiva il Foro di Mantova."
    End If

    ' SEPA
    f("RIF_MANDATO_SEPA") = "SDD-" & NomeSicuro(f("CODICE_CLIENTE")) & "-" & Format$(CDateSicura(f("DATA_LETTERA")), "yyyymmdd")
    If f("INTESTATARIO_CONTO") = "" Then f("INTESTATARIO_CONTO") = f("DENOMINAZIONE")
    If f("CF_INTESTATARIO") = "" Then f("CF_INTESTATARIO") = IIf(f("PIVA") <> ChrW(8212), f("PIVA"), f("CF"))
    If f("BIC_CLIENTE") = "" Then f("BIC_CLIENTE") = ChrW(8212)
    f("SOTTOSCRITTORE_SEPA") = IIf(isSoc And lr <> "", lr, ChrW(8212))

    ' termine di restituzione
    gg = 15
    If IsNumeric(f("GG_TERMINE_RESTITUZIONE")) And f("GG_TERMINE_RESTITUZIONE") <> "" Then gg = CLng(f("GG_TERMINE_RESTITUZIONE"))
    f("TERMINE_RESTITUZIONE") = Format$(CDateSicura(f("DATA_LETTERA")) + gg, "dd/mm/yyyy")

    ' condizioni (paragrafi / righe con marcatore {{?FLAG}})
    For Each p In Split(PRESTAZIONI, ",")
        flags("P_" & p) = SiNo(f("P_" & p))
    Next p
    flags("COMMITTENTE_FORTE") = SiNo(f("COMMITTENTE_FORTE"))
    flags("NON_CONSUMATORE") = Not isCons
    flags("CLAUSOLA_ART28") = SiNo(f("CLAUSOLA_ART28"))
    flags("SPESE_GENERALI") = (f("SPESE_GENERALI_PERC") <> "")
    flags("SDD") = (UCase$(f("PAGAMENTO")) = "SDD")

    ' totale annuo delle prestazioni periodiche (importo x periodi annui)
    totale = 0
    For Each p In Split(PRESTAZIONI, ",")
        If flags("P_" & p) And IsNumeric(f("RAW_IMP_" & p)) And Not IsEmpty(f("RAW_IMP_" & p)) Then
            Select Case LCase$(f("PER_" & p))
                Case "mensile": molt = 12
                Case "trimestrale": molt = 4
                Case "semestrale": molt = 2
                Case "annuale": molt = 1
                Case Else: molt = 0
            End Select
            totale = totale + CDbl(f("RAW_IMP_" & p)) * molt
        End If
    Next p
    f("TOTALE_ANNUO") = Format$(totale, "#,##0.00")
End Sub

Private Function CDateSicura(ByVal s As String) As Date
    If IsDate(s) Then CDateSicura = CDate(s) Else CDateSicura = Date
End Function

'==============================================================================
' VALIDAZIONE (errori bloccanti)
'==============================================================================
Private Function Valida(ByVal f As Object, ByVal flags As Object) As String
    Dim e As String, k As Variant, p As Variant, nPrest As Long, tipo As String

    For Each k In Array("CODICE_CLIENTE", "DENOMINAZIONE", "TIPO_CLIENTE", "INDIRIZZO", "CAP", "CITTA", "CF", _
                        "EMAIL_DESTINATARIO", "DURATA", "GG_DISDETTA", "GG_CONSEGNA_DOC", "DATA_CONSEGNA_DICH", _
                        "COMPLESSITA", "MOTIVAZIONE_COMPLESSITA", "FATTURAZIONE", "GG_PAGAMENTO", "N_FATTURE_RECESSO", _
                        "AGGIORNAMENTO_COMPENSI", "DATA_COMUNICAZIONE_AGG", "PAGAMENTO")
        If f(k) = "" Then e = e & "; " & k & " mancante"
    Next k
    If Not IsDate(f("DATA_LETTERA")) Then e = e & "; DATA_LETTERA non valida"

    tipo = LCase$(f("TIPO_CLIENTE"))
    If InStr(tipo, "societ") > 0 And f("LEGALE_RAPP") = "" Then e = e & "; LEGALE_RAPP obbligatorio per le societ" & ChrW(224)
    If Not EmailValide(f("EMAIL_DESTINATARIO")) Then e = e & "; EMAIL_DESTINATARIO non valida"
    If f("EMAIL_CC") <> "" And Not EmailValide(f("EMAIL_CC")) Then e = e & "; EMAIL_CC non valida"

    ' prestazioni e compensi (art. 9, c. 4, D.L. 1/2012: preventivo per ogni prestazione)
    For Each p In Split(PRESTAZIONI, ",")
        If flags("P_" & p) Then
            nPrest = nPrest + 1
            If f("MOD_" & p) = "" Then e = e & "; MOD_" & p & " mancante"
            If f("IMP_" & p) = "" Then e = e & "; IMP_" & p & " mancante"
            If f("PER_" & p) = "" Then e = e & "; PER_" & p & " mancante"
        End If
    Next p
    If nPrest = 0 Then e = e & "; nessuna prestazione conferita (colonne P_)"
    If flags("P_CONT") And f("REGIME_CONTABILE") = "" Then e = e & "; REGIME_CONTABILE mancante"
    If flags("P_BIL") And f("SCHEMA_BILANCIO") = "" Then e = e & "; SCHEMA_BILANCIO mancante"
    If flags("P_STRA") And f("DESCR_STRA") = "" Then e = e & "; DESCR_STRA mancante"
    If flags("P_PER") And f("DESCR_PER") = "" Then e = e & "; DESCR_PER mancante"
    If flags("P_ALT") And f("DESCR_ALT") = "" Then e = e & "; DESCR_ALT mancante"

    ' indipendenza del revisore (art. 10 D.Lgs. 39/2010)
    If SiNo(f("CLIENTE_REVISIONE")) And (flags("P_CONT") Or flags("P_BIL")) Then
        e = e & "; INDIPENDENZA: cliente soggetto a revisione legale dello Studio - tenuta contabilit" & ChrW(224) & _
                " / bilancio incompatibili (art. 10 D.Lgs. 39/2010)"
    End If
    ' equo compenso (L. 49/2023)
    If flags("COMMITTENTE_FORTE") And Not SiNo(f("VERIFICA_EQUO_COMPENSO")) Then
        e = e & "; EQUO COMPENSO: committente forte - confermare in VERIFICA_EQUO_COMPENSO il rispetto dei parametri D.M. 140/2012"
    End If

    ' pagamento
    Select Case UCase$(f("PAGAMENTO"))
        Case "BONIFICO"
            If Not IbanValido(f("IBAN_STUDIO")) Then e = e & "; Impostazioni: IBAN_STUDIO non valido"
        Case "SDD"
            If Not IbanValido(f("IBAN_CLIENTE")) Then e = e & "; IBAN_CLIENTE non valido"
            If Impostazione("SEPA_CID") = "" Or InStr(Impostazione("SEPA_CID"), "[DC]") > 0 Then e = e & "; Impostazioni: SEPA_CID mancante"
        Case ""
        Case Else
            e = e & "; PAGAMENTO deve essere Bonifico o SDD"
    End Select

    ' impostazioni obbligatorie dello Studio (polizza RC: art. 9, c. 4, D.L. 1/2012; art. 5 DPR 137/2012)
    For Each k In Array("ACCOUNT_MITTENTE", "CASSA_PREVIDENZA", "TARIFFA_TITOLARE", "TARIFFA_COLLABORATORI", _
                        "RC_COMPAGNIA", "RC_POLIZZA", "RC_MASS_SINISTRO", "RC_MASS_ANNUO", "RC_SCADENZA", "RC_SINDACO_REVISORE")
        If Impostazione(CStr(k)) = "" Or InStr(Impostazione(CStr(k)), "[DC]") > 0 Then e = e & "; Impostazioni: " & k & " da compilare"
    Next k

    ' modelli
    For Each k In Array("MODELLO_MANDATO", "MODELLO_AML", "MODELLO_SEPA", "MODELLO_PRIVACY")
        If Dir$(PercorsoModello(CStr(k))) = "" Then e = e & "; modello non trovato: " & PercorsoModello(CStr(k))
    Next k

    If Len(e) > 0 Then e = Mid$(e, 3)
    Valida = e
End Function

Private Function EmailValide(ByVal s As String) As Boolean
    Dim a As Variant, x As String
    If Trim$(s) = "" Then Exit Function
    For Each a In Split(Replace(s, ",", ";"), ";")
        x = Trim$(CStr(a))
        If x <> "" Then
            If Not (x Like "?*@?*.?*") Or InStr(x, " ") > 0 Then Exit Function
        End If
    Next a
    EmailValide = True
End Function

Private Function IbanValido(ByVal s As String) As Boolean
    s = UCase$(Replace(s, " ", ""))
    If Len(s) < 15 Or Len(s) > 34 Then Exit Function
    If Not (Left$(s, 2) Like "[A-Z][A-Z]") Then Exit Function
    If Left$(s, 2) = "IT" And Len(s) <> 27 Then Exit Function
    IbanValido = True
End Function

'==============================================================================
' WORD: GENERAZIONE DOCUMENTI
'==============================================================================
Private Function GeneraDocumento(ByVal modello As String, ByVal f As Object, ByVal flags As Object, _
                                 ByVal cartella As String, ByVal nomeBase As String, ByRef errMsg As String) As String
    Dim doc As Object, k As Variant, pdf As String, residuo As String

    errMsg = ""
    If gWord Is Nothing Then
        Set gWord = CreateObject("Word.Application")
        gWord.Visible = False
        gWord.DisplayAlerts = 0
    End If
    Set doc = gWord.Documents.Add(Template:=modello, Visible:=False)

    ApplicaCondizioni doc, flags
    For Each k In f.Keys
        If Left$(k, 4) <> "RAW_" Then SostituisciOvunque doc, "{{" & k & "}}", CStr(f(k))
    Next k

    residuo = PrimoResiduo(doc)
    If residuo <> "" Then
        errMsg = "Campo non compilato in " & nomeBase & ": " & residuo & " (aggiungere la colonna nel foglio Clienti o la voce in Impostazioni)"
        doc.Close wdDoNotSaveChanges
        Exit Function
    End If
    If TrovaTesto(doc, "[DC]") Then
        errMsg = "Dato [DC] (da completare) presente in " & nomeBase & ": completare le Impostazioni"
        doc.Close wdDoNotSaveChanges
        Exit Function
    End If

    TogliEvidenziazione doc
    doc.Fields.Update
    If SiNo(Impostazione("SALVA_DOCX")) Then
        doc.SaveAs2 Filename:=cartella & "\" & nomeBase & ".docx", FileFormat:=wdFormatXMLDocument
    End If
    pdf = cartella & "\" & nomeBase & ".pdf"
    doc.ExportAsFixedFormat OutputFileName:=pdf, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
                            OptimizeFor:=wdExportOptimizeForPrint, UseISO19005_1:=SiNo(Impostazione("PDF_A"))
    doc.Close wdDoNotSaveChanges
    GeneraDocumento = pdf
End Function

' Marcatori {{?FLAG}}: se FLAG vero si elimina il marcatore, altrimenti si elimina il paragrafo
' (o la riga di tabella) che lo contiene.
Private Sub ApplicaCondizioni(ByVal doc As Object, ByVal flags As Object)
    Dim k As Variant, rng As Object, marker As String, guard As Long
    For Each k In flags.Keys
        marker = "{{?" & k & "}}"
        guard = 0
        Do
            Set rng = doc.Content
            With rng.Find
                .ClearFormatting
                .Text = marker
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = True
                .MatchWildcards = False
                .Format = False
            End With
            If Not rng.Find.Execute Then Exit Do
            If flags(k) Then
                rng.Text = ""
            ElseIf rng.Information(wdWithInTable) Then
                rng.Rows(1).Delete
            Else
                rng.Paragraphs(1).Range.Delete
            End If
            guard = guard + 1
            If guard > 200 Then Exit Do
        Loop
    Next k
End Sub

Private Sub SostituisciOvunque(ByVal doc As Object, ByVal cerca As String, ByVal sostituisci As String)
    Dim story As Object, rng As Object
    For Each story In doc.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            SostituisciInRange rng, cerca, sostituisci
            Set rng = rng.NextStoryRange
        Loop
    Next story
End Sub

' Sostituzione tramite Range.Text: nessun limite di 255 caratteri.
Private Sub SostituisciInRange(ByVal story As Object, ByVal cerca As String, ByVal sostituisci As String)
    Dim rng As Object, guard As Long
    Set rng = story.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = cerca
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = False
        .Format = False
    End With
    Do While rng.Find.Execute
        rng.Text = sostituisci
        rng.HighlightColorIndex = wdNoHighlight
        rng.Collapse wdCollapseEnd
        guard = guard + 1
        If guard > 500 Then Exit Do
    Loop
End Sub

Private Function TrovaTesto(ByVal doc As Object, ByVal testo As String) As Boolean
    Dim story As Object, rng As Object
    For Each story In doc.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            If InStr(1, rng.Text, testo, vbBinaryCompare) > 0 Then TrovaTesto = True: Exit Function
            Set rng = rng.NextStoryRange
        Loop
    Next story
End Function

Private Function PrimoResiduo(ByVal doc As Object) As String
    Dim story As Object, rng As Object, t As String, i As Long, j As Long
    For Each story In doc.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            t = rng.Text
            i = InStr(1, t, "{{")
            If i > 0 Then
                j = InStr(i, t, "}}")
                If j > 0 Then PrimoResiduo = Mid$(t, i, j - i + 2) Else PrimoResiduo = Mid$(t, i, 40)
                Exit Function
            End If
            Set rng = rng.NextStoryRange
        Loop
    Next story
End Function

Private Sub TogliEvidenziazione(ByVal doc As Object)
    Dim story As Object, rng As Object
    For Each story In doc.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            rng.HighlightColorIndex = wdNoHighlight
            Set rng = rng.NextStoryRange
        Loop
    Next story
End Sub

Private Sub ChiudiWord()
    On Error Resume Next
    If Not gWord Is Nothing Then
        Dim d As Object
        For Each d In gWord.Documents
            d.Close wdDoNotSaveChanges
        Next d
        gWord.Quit wdDoNotSaveChanges
    End If
    Set gWord = Nothing
End Sub

'==============================================================================
' OUTLOOK: INVIO
'==============================================================================
Private Function InviaEmail(ByVal f As Object, ByVal allegati As Collection, ByVal descr As Collection, _
                            ByVal modalita As String) As String
    Dim mail As Object, acc As Object, a As Variant, oggetto As String

    On Error GoTo ErroreMail
    If gOutlook Is Nothing Then Set gOutlook = CreateObject("Outlook.Application")
    Set acc = TrovaAccount(Impostazione("ACCOUNT_MITTENTE"))
    If acc Is Nothing Then
        InviaEmail = "ERRORE: account " & Impostazione("ACCOUNT_MITTENTE") & " non configurato in Outlook"
        Exit Function
    End If

    oggetto = CompilaTesto(TestoCover("OGGETTO"), f, False)
    Set mail = gOutlook.CreateItem(olMailItem)
    Set mail.SendUsingAccount = acc
    mail.To = f("EMAIL_DESTINATARIO")
    If f("EMAIL_CC") <> "" Then mail.CC = f("EMAIL_CC")
    If Impostazione("CCN_ARCHIVIO") <> "" Then mail.BCC = Impostazione("CCN_ARCHIVIO")
    mail.Subject = oggetto
    mail.HTMLBody = CorpoHtml(f, descr)
    For Each a In allegati
        mail.Attachments.Add CStr(a)
    Next a

    Select Case modalita
        Case "INVIO": mail.Send
        Case "BOZZA": mail.Save
        Case Else: mail.Display
    End Select
    InviaEmail = modalita & " " & Format$(Now, "dd/mm/yyyy hh:nn") & " a " & f("EMAIL_DESTINATARIO") & _
                 " - " & allegati.Count & " allegati"
    Exit Function
ErroreMail:
    InviaEmail = "ERRORE invio: " & Err.Description
End Function

Private Function TrovaAccount(ByVal smtp As String) As Object
    Dim a As Object
    For Each a In gOutlook.Session.Accounts
        If LCase$(a.SmtpAddress) = LCase$(Trim$(smtp)) Then Set TrovaAccount = a: Exit Function
    Next a
End Function

Private Function TestoCover(ByVal chiave As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_COVER)
    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = chiave Then TestoCover = CStr(ws.Cells(r, 2).Value): Exit Function
    Next r
End Function

Private Function CompilaTesto(ByVal t As String, ByVal f As Object, ByVal html As Boolean) As String
    Dim k As Variant
    For Each k In f.Keys
        If Left$(k, 4) <> "RAW_" Then
            If html Then
                t = Replace(t, "{{" & k & "}}", HtmlEsc(CStr(f(k))), , , vbTextCompare)
            Else
                t = Replace(t, "{{" & k & "}}", CStr(f(k)), , , vbTextCompare)
            End If
        End If
    Next k
    CompilaTesto = t
End Function

Private Function CorpoHtml(ByVal f As Object, ByVal descr As Collection) As String
    Dim t As String, lista As String, rest As String, d As Variant, isSoc As Boolean

    For Each d In descr
        lista = lista & "<li>" & HtmlEsc(CStr(d)) & "</li>"
    Next d
    lista = "<ul style=""margin-top:0"">" & lista & "</ul>"

    isSoc = (InStr(LCase$(f("TIPO_CLIENTE")), "societ") > 0)
    rest = "<li>la lettera di incarico sottoscritta in calce e nella sezione di approvazione specifica delle clausole ai sensi degli artt. 1341 e 1342 c.c.;</li>"
    If SiNo(f("ALL_AML")) Then
        rest = rest & "<li>la dichiarazione antiriciclaggio (M3) compilata e sottoscritta, con la presa visione dell'informativa (M1);</li>" & _
               "<li>copia del documento di identit&agrave; in corso di validit&agrave; del sottoscrittore" & _
               IIf(isSoc, " e dei titolari effettivi, e visura camerale aggiornata", "") & ";</li>"
    End If
    If UCase$(f("PAGAMENTO")) = "SDD" Then rest = rest & "<li>il mandato SEPA sottoscritto dall'intestatario del conto;</li>"
    If SiNo(f("ALL_PRIVACY")) Then rest = rest & "<li>l'informativa privacy sottoscritta per presa visione.</li>"
    rest = "<ul style=""margin-top:0"">" & rest & "</ul>"

    t = HtmlEsc(TestoCover("CORPO"))
    t = CompilaTesto(t, f, True)
    t = Replace(t, "{{ELENCO_ALLEGATI}}", lista)
    t = Replace(t, "{{ELENCO_RESTITUZIONE}}", rest)
    t = Replace(t, vbCr, "")
    t = Replace(t, "</ul>" & vbLf, "</ul>")
    t = Replace(t, vbLf, "<br>")

    CorpoHtml = "<html><body><div style=""font-family:Arial,sans-serif;font-size:10pt;color:#000"">" & t & _
                "<br><br><span style=""font-size:8pt;color:#777"">" & HtmlEsc(TestoCover("DISCLAIMER")) & "</span></div></body></html>"
End Function

Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEsc = s
End Function

'==============================================================================
' FILE, ESITI, LOG
'==============================================================================
Private Function CartellaBase() As String
    CartellaBase = ThisWorkbook.Path
    If LCase$(Left$(CartellaBase, 4)) = "http" Then
        Err.Raise vbObjectError + 2, , "Il file Excel " & ChrW(232) & " aperto da OneDrive/SharePoint (percorso web): " & _
                  "indicare CARTELLA_MODELLI e CARTELLA_OUTPUT come percorsi locali nel foglio Impostazioni."
    End If
End Function

Private Function PercorsoModello(ByVal chiave As String) As String
    Dim base As String
    base = Impostazione("CARTELLA_MODELLI")
    If base = "" Then base = CartellaBase()
    If Right$(base, 1) = "\" Then base = Left$(base, Len(base) - 1)
    PercorsoModello = base & "\" & Impostazione(chiave)
End Function

Private Function CartellaOutput() As String
    Dim base As String
    base = Impostazione("CARTELLA_OUTPUT")
    If base = "" Then base = CartellaBase() & "\Mandati_generati"
    If Right$(base, 1) = "\" Then base = Left$(base, Len(base) - 1)
    CartellaOutput = base
End Function

Private Sub CreaCartella(ByVal percorso As String)
    Dim fso As Object, padre As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(percorso) Then Exit Sub
    padre = fso.GetParentFolderName(percorso)
    If Len(padre) > 0 And Not fso.FolderExists(padre) Then CreaCartella padre
    fso.CreateFolder percorso
End Sub

Private Function NomeSicuro(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "_")
    Next c
    NomeSicuro = Trim$(s)
End Function

Private Function ElencoFile(ByVal c As Collection) As String
    Dim x As Variant, s As String
    For Each x In c
        s = s & " | " & Mid$(CStr(x), InStrRev(CStr(x), "\") + 1)
    Next x
    If Len(s) > 0 Then s = Mid$(s, 4)
    ElencoFile = s
End Function

Private Sub ScriviEsito(ByVal ws As Worksheet, ByVal r As Long, ByVal stato As String, _
                        ByVal esito As String, ByVal cartella As String)
    If stato <> "" Then ws.Cells(r, Col("STATO")).Value = stato
    ws.Cells(r, Col("ESITO")).Value = esito
    If cartella <> "" Then ws.Cells(r, Col("CARTELLA_DOCUMENTI")).Value = cartella
End Sub

Private Sub ScriviLog(ByVal f As Object, ByVal fase As String, ByVal esito As String, _
                     ByVal dettaglio As String, ByVal files As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = Environ$("USERNAME")
    ws.Cells(r, 3).Value = f("CODICE_CLIENTE")
    ws.Cells(r, 4).Value = f("DENOMINAZIONE")
    ws.Cells(r, 5).Value = f("EMAIL_DESTINATARIO")
    ws.Cells(r, 6).Value = fase
    ws.Cells(r, 7).Value = esito
    ws.Cells(r, 8).Value = dettaglio
    ws.Cells(r, 9).Value = files
End Sub
