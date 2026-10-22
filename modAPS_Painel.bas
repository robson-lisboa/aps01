Attribute VB_Name = "modAPS_Painel"
Option Explicit

'==========================================================
' APS PURAN - PAINEL (dashboard) - versao simplificada em 3 niveis
'
'   VISAO GERAL  -> como esta o planejamento, o que precisa de atencao, onde decidir
'   MAQUINAS     -> comparacao entre maquinas (e acesso ao detalhe de cada uma)
'   DETALHES     -> uma maquina (ordens do periodo) OU a lista de ordens que precisam de atencao
'
' Tudo acontece na aba "Painel". A navegacao e feita por botoes (Voltar, Visao geral, Maquinas, Ver detalhes);
' a tela atual fica guardada em celulas ocultas (coluna Q), por isso o filtro de periodo e a maquina escolhida
' sobrevivem ao "Voltar".
'
' O painel SOMENTE LE 02_Operacoes (via APS_LerOps) e o calendario existente.
' NAO grava em 02_Operacoes, 1_Planejamento nem no calendario; NAO altera cards, motor, formularios ou diagnostico.
'
' Regras de calculo (inalteradas):
'  * Comprimidos = Caixas x APS_CP_CAIXA (30). E quantidade PLANEJADA; nao existe registro de producao real.
'  * Operacao no periodo = a que se SOBREPOE ao periodo [inicio, fim+1 dia). Cancelada nao conta.
'  * Quantidade e horas de uma operacao que atravessa o periodo sao proporcionais ao TEMPO DE TRABALHO (calendario).
'  * Ocupacao = (horas de producao + Setup + Limpeza) / horas de trabalho do calendario no periodo.
'  * Setup/Limpeza: um bloco por producao (montado por APS_LerOps a partir da coluna Setup); sem dupla contagem.
'  * Sem dado real (nao exibidos): produzido, tempo realizado, disponibilidade, desempenho, qualidade, OEE.
'==========================================================

Private Const PN_ABA As String = "Painel"
Private Const PN_LIN0 As Long = 8            ' primeira linha de conteudo (linhas 1 a 7 = cabecalho fixo)
Private Const PN_COLN As Long = 11           ' ultima coluna visivel (K)
Private Const PN_COLEST As Long = 17         ' coluna Q (oculta): estado da navegacao
Private Const PN_COLLST As Long = 18         ' coluna R (oculta): listas das caixas de selecao
Private Const PN_COLGRAF As Long = 19        ' coluna S (oculta): dados dos graficos
Private Const PN_VERSAO As String = "v2"
Private Const PN_MAXDET As Long = 400
Private Const PN_MAXPRI As Long = 30
Private Const PN_T_GERAL As String = "geral"
Private Const PN_T_MAQ As String = "maquinas"
Private Const PN_T_DET As String = "detalhe"
Private Const PN_M_MAQ As String = "maquina"
Private Const PN_M_ATENCAO As String = "atencao"
Private Const PN_CPER As String = "C5"       ' filtro: periodo
Private Const PN_CREF As String = "D5"       ' filtro: data de referencia
Private Const PN_CMAQ As String = "F5"       ' filtro: maquina (so na tela de detalhe da maquina)

Public APS_PainelOcupado As Boolean
Public APS_PainelUltAssin As Double

Private Type tAgg
    ordens As Long
    nPlan As Long
    nAnd As Long
    nConc As Long
    nAtr As Long
    nConf As Long
    nCanc As Long
    nAtrReg As Long
    nAtrAny As Long
    nSemQtd As Long
    comp As Double
    hProd As Double
    hSetup As Double
    hLimp As Double
    nSetup As Long
    nLimp As Long
    hAtrReg As Double
End Type

Private Type tPri
    pri As Long
    tipo As String
    tit As String
    maq As String
    ini As Double
    fim As Double
    txt As String
End Type

Private pnMaq() As String
Private pnNM As Long
Private pnOps() As tOperacao
Private pnN As Long
Private pnConfP() As Boolean
Private pnAl() As tAlerta
Private pnNAl As Long
Private pnFalhaGraf As String
Private pnSel As Long          ' indice da maquina selecionada (0 = todas)
Private pnPIni As Double, pnPFim As Double, pnAIni As Double, pnAFim As Double
Private pnDispC As Double, pnDispP As Double
Private pnAgC() As tAgg
Private pnAgP() As tAgg
Private pnRot As String, pnRotA As String
Private pnNB As Long
Private pnBI() As Double, pnBF() As Double, pnBL() As String
Private pnBComp() As Double, pnBOc() As Double, pnBDisp() As Double
Private pnPri() As tPri
Private pnNPri As Long
Private pnPasso As String

'----------------------------------------------------------
' Cores (paleta reduzida: azul, cinzas e vermelho so para alerta)
'----------------------------------------------------------
Private Function CorAz() As Long
    CorAz = RGB(31, 78, 120)
End Function
Private Function CorCinza() As Long
    CorCinza = RGB(115, 115, 115)
End Function
Private Function CorTxt() As Long
    CorTxt = RGB(38, 50, 56)
End Function
Private Function CorAlerta() As Long
    CorAlerta = RGB(192, 0, 0)
End Function
Private Function CorCartao() As Long
    CorCartao = RGB(242, 244, 246)
End Function
Private Function CorLinha() As Long
    CorLinha = RGB(222, 226, 230)
End Function

'----------------------------------------------------------
' Pontos de entrada
'----------------------------------------------------------
' Abre o painel na VISAO GERAL (tela inicial)
Public Sub APS_PainelAbrir()
    Dim ws As Worksheet
    Set ws = PN_AbaPainel(True)
    If ws Is Nothing Then Exit Sub
    PN_SetEst ws, PN_T_GERAL, "", "", ""
    APS_PainelAtualizar False
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A7").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Botao "Atualizar": recalcula a tela em que o usuario esta, com os filtros escolhidos
Public Sub APS_PainelAtualizarBotao()
    APS_PainelAtualizar False
End Sub

' Chamada automatica (de APS_Atualizar): so atualiza se a aba existe e os dados de 02_Operacoes mudaram.
Public Sub APS_PainelAuto()
    Dim ws As Worksheet, a As Double
    If APS_PainelOcupado Then Exit Sub
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    a = APS_AssinaturaOps()
    On Error GoTo 0
    If a <> APS_PainelUltAssin Then APS_PainelAtualizar True
End Sub

' OPCIONAL: se o usuario quiser que o painel atualize sozinho ao mudar um filtro, ThisWorkbook chama esta rotina
' a partir de Workbook_SheetChange (ver LEIA-ME). Sem isso, basta clicar em Atualizar.
Public Sub APS_PainelAoAlterar(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet
    If APS_PainelOcupado Then Exit Sub
    If StrComp(Sh.Name, PN_ABA, vbTextCompare) <> 0 Then Exit Sub
    Set ws = Sh
    If Intersect(Target, ws.Range("C5:D5,F5")) Is Nothing Then Exit Sub
    APS_PainelAtualizar True
End Sub

Public Sub APS_PainelAtualizar(Optional ByVal silencioso As Boolean = False)
    Dim ws As Worksheet, ev As Boolean, su As Boolean, msg As String

    If APS_PainelOcupado Then Exit Sub
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then
        If Not silencioso Then MsgBox "A aba '" & PN_ABA & "' ainda n" & ChrW(227) & "o existe. Execute APS_PainelAbrir.", vbInformation, APS_TITULO
        Exit Sub
    End If
    If ws.ProtectContents Then
        If Not silencioso Then MsgBox "A aba '" & PN_ABA & "' est" & ChrW(225) & " protegida. Desproteja-a para atualizar o painel.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    APS_PainelOcupado = True
    ev = Application.EnableEvents
    su = Application.ScreenUpdating
    On Error GoTo Falha
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    PN_Montar ws
    APS_PainelUltAssin = APS_AssinaturaOps()
    GoTo Sai

Falha:
    msg = "Falha ao atualizar o painel: " & Err.Description & " [erro " & Err.Number & " | etapa: " & pnPasso & " | blocos=" & pnN & " maquinas=" & pnNM & "]"
Sai:
    On Error Resume Next
    Application.EnableEvents = ev
    Application.ScreenUpdating = su
    APS_PainelOcupado = False
    If Len(msg) > 0 Then
        ws.Range("B6").Value = msg
        ws.Range("B6").Font.Color = CorAlerta()
        If Not silencioso Then MsgBox msg, vbExclamation, APS_TITULO
    End If
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' Navegacao (rotinas ligadas aos botoes da aba)
'----------------------------------------------------------
Public Sub APS_PainelNavGeral()
    PN_Ir PN_T_GERAL, "", "", ""
End Sub

Public Sub APS_PainelNavMaquinas()
    PN_Ir PN_T_MAQ, PN_T_GERAL, "", ""
End Sub

' Lista completa de ordens que precisam de atencao (Detalhes)
Public Sub APS_PainelNavAtencao()
    Dim ws As Worksheet, cur As String
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then Exit Sub
    cur = PN_Est(ws, 1)
    If StrComp(cur, PN_T_MAQ, vbTextCompare) <> 0 Then cur = PN_T_GERAL
    PN_Ir PN_T_DET, cur, "", PN_M_ATENCAO
End Sub

' Detalhe de uma maquina: o botao clicado guarda o nome da maquina no texto alternativo
Public Sub APS_PainelNavMaquina()
    Dim ws As Worksheet, nm As String, cur As String
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    nm = ws.Shapes(CStr(Application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(nm) = 0 Then Exit Sub
    cur = PN_Est(ws, 1)
    If StrComp(cur, PN_T_MAQ, vbTextCompare) <> 0 Then cur = PN_T_GERAL
    PN_Ir PN_T_DET, cur, nm, PN_M_MAQ
End Sub

' Voltar: sempre para a tela de onde o usuario veio (padrao: Visao geral)
Public Sub APS_PainelNavVoltar()
    Dim ws As Worksheet, t As String, ant As String
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then Exit Sub
    t = PN_Est(ws, 2)
    If StrComp(t, PN_T_MAQ, vbTextCompare) = 0 Then
        ant = PN_T_GERAL
    Else
        t = PN_T_GERAL
        ant = ""
    End If
    PN_Ir t, ant, "", ""
End Sub

Private Sub PN_Ir(ByVal tela As String, ByVal ant As String, ByVal maq As String, ByVal modo As String)
    Dim ws As Worksheet
    If APS_PainelOcupado Then Exit Sub
    Set ws = PN_AbaPainel(False)
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then
        MsgBox "A aba '" & PN_ABA & "' est" & ChrW(225) & " protegida. Desproteja-a para navegar.", vbExclamation, APS_TITULO
        Exit Sub
    End If
    PN_SetEst ws, tela, ant, maq, modo
    If StrComp(tela, PN_T_DET, vbTextCompare) = 0 And StrComp(modo, PN_M_MAQ, vbTextCompare) = 0 And Len(maq) > 0 Then
        ws.Range(PN_CMAQ).Value = maq
    End If
    APS_PainelAtualizar False
    On Error Resume Next
    If ActiveWindow.FreezePanes Then ActiveWindow.ScrollRow = 7 Else ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Estado da navegacao: Q1 = tela, Q2 = tela anterior (destino do Voltar), Q3 = maquina, Q4 = tipo de detalhe, Q5 = versao do layout
Private Function PN_Est(ByVal ws As Worksheet, ByVal i As Long) As String
    PN_Est = Trim$(CStr(ws.Cells(i, PN_COLEST).Value))
End Function

Private Sub PN_SetEst(ByVal ws As Worksheet, ByVal tela As String, ByVal ant As String, ByVal maq As String, ByVal modo As String)
    On Error Resume Next
    ws.Cells(1, PN_COLEST).Value = tela
    ws.Cells(2, PN_COLEST).Value = ant
    If Len(maq) > 0 Then ws.Cells(3, PN_COLEST).Value = maq
    ws.Cells(4, PN_COLEST).Value = modo
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' Aba (cria so quando pedido; converte o layout antigo na primeira vez)
'----------------------------------------------------------
Private Function PN_AbaPainel(ByVal criar As Boolean) As Worksheet
    Dim ws As Worksheet
    pnPasso = "PN_AbaPainel"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PN_ABA)
    On Error GoTo 0
    If ws Is Nothing And criar Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PN_ABA
        PN_Estrutura ws
    ElseIf Not ws Is Nothing Then
        If Not ws.ProtectContents Then
            If StrComp(PN_Est(ws, 5), PN_VERSAO, vbTextCompare) <> 0 Then PN_Estrutura ws
        End If
    End If
    Set PN_AbaPainel = ws
End Function

' Prepara a aba (nova ou do layout antigo): larguras, colunas ocultas, estado e valores iniciais dos filtros.
' No layout antigo o periodo estava em D4 e a data em D5: esses dois valores sao aproveitados.
Private Sub PN_Estrutura(ByVal ws As Worksheet)
    Dim per As String, ref As Variant, i As Long

    pnPasso = "PN_Estrutura"
    If Len(PN_Est(ws, 5)) = 0 Then
        per = Trim$(CStr(ws.Range("D4").Value))
        ref = ws.Range("D5").Value
    Else
        per = Trim$(CStr(ws.Range(PN_CPER).Value))
        ref = ws.Range(PN_CREF).Value
    End If

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "pnl_" Then ws.Shapes(i).Delete
    Next i
    On Error Resume Next
    ws.Cells.Hyperlinks.Delete
    ws.Cells.UnMerge
    ws.Cells.Validation.Delete
    ws.Cells.Clear
    ws.Cells.RowHeight = ws.StandardHeight
    ws.Columns("A:Z").Hidden = False
    On Error GoTo 0

    With ws.Cells
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CorTxt()
        .Interior.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 42
    For i = 3 To PN_COLN
        ws.Columns(i).ColumnWidth = 14
    Next i
    ws.Columns(PN_COLN + 1).ColumnWidth = 2
    ws.Range(ws.Columns(PN_COLEST), ws.Columns(26)).Hidden = True

    ws.Range(ws.Cells(1, PN_COLEST), ws.Cells(5, PN_COLEST)).NumberFormat = "@"
    ws.Cells(1, PN_COLEST).Value = PN_T_GERAL
    ws.Cells(3, PN_COLEST).Value = "Todas"
    ws.Cells(5, PN_COLEST).Value = PN_VERSAO

    ws.Cells(1, PN_COLLST).Value = "Semanal"
    ws.Cells(2, PN_COLLST).Value = "Mensal"
    ws.Cells(3, PN_COLLST).Value = "Semestral"
    ws.Cells(4, PN_COLLST).Value = "Anual"

    If StrComp(per, "Semanal", vbTextCompare) <> 0 And StrComp(per, "Mensal", vbTextCompare) <> 0 And _
       StrComp(per, "Semestral", vbTextCompare) <> 0 And StrComp(per, "Anual", vbTextCompare) <> 0 Then per = "Mensal"
    ws.Range(PN_CPER).Value = per
    ws.Range(PN_CREF).NumberFormat = "dd/mm/yyyy"
    If IsDate(ref) Then ws.Range(PN_CREF).Value = CDate(ref) Else ws.Range(PN_CREF).Value = Date
End Sub

'----------------------------------------------------------
' Periodo selecionado e periodo anterior equivalente. pFim e EXCLUSIVO (dia seguinte, 00:00).
'----------------------------------------------------------
Private Function PN_Periodo(ByVal ws As Worksheet, ByRef pIni As Double, ByRef pFim As Double, _
                            ByRef aIni As Double, ByRef aFim As Double, ByRef rot As String, ByRef rotA As String, _
                            ByRef msg As String) As Boolean
    Dim tp As String, ref As Date, d As Date, y As Long, m As Long

    pnPasso = "PN_Periodo"
    tp = Trim$(CStr(ws.Range(PN_CPER).Value))
    If StrComp(tp, "Semanal", vbTextCompare) <> 0 And StrComp(tp, "Mensal", vbTextCompare) <> 0 And _
       StrComp(tp, "Semestral", vbTextCompare) <> 0 And StrComp(tp, "Anual", vbTextCompare) <> 0 Then
        msg = "Escolha o per" & ChrW(237) & "odo em " & PN_CPER & " (Semanal, Mensal, Semestral ou Anual) e clique em Atualizar."
        Exit Function
    End If

    If IsDate(ws.Range(PN_CREF).Value) Then ref = CDate(ws.Range(PN_CREF).Value) Else ref = Date
    d = Int(ref)
    y = Year(d): m = Month(d)

    Select Case LCase$(tp)
        Case "semanal"
            pIni = CDbl(d) - (Weekday(d, vbMonday) - 1)
            pFim = pIni + 7
            aIni = pIni - 7: aFim = pIni
            rot = "semana de " & Format$(pIni, "dd/mm/yyyy") & " a " & Format$(pFim - 1, "dd/mm/yyyy")
        Case "mensal"
            pIni = CDbl(DateSerial(y, m, 1)): pFim = CDbl(DateSerial(y, m + 1, 1))
            aIni = CDbl(DateSerial(y, m - 1, 1)): aFim = pIni
            rot = "m" & ChrW(234) & "s " & Format$(pIni, "mm/yyyy") & " (" & Format$(pIni, "dd/mm") & " a " & Format$(pFim - 1, "dd/mm/yyyy") & ")"
        Case "semestral"
            If m <= 6 Then
                pIni = CDbl(DateSerial(y, 1, 1)): pFim = CDbl(DateSerial(y, 7, 1))
                aIni = CDbl(DateSerial(y - 1, 7, 1)): aFim = pIni
                rot = "1" & ChrW(186) & " semestre de " & y & " (01/01 a 30/06)"
            Else
                pIni = CDbl(DateSerial(y, 7, 1)): pFim = CDbl(DateSerial(y + 1, 1, 1))
                aIni = CDbl(DateSerial(y, 1, 1)): aFim = pIni
                rot = "2" & ChrW(186) & " semestre de " & y & " (01/07 a 31/12)"
            End If
        Case Else
            pIni = CDbl(DateSerial(y, 1, 1)): pFim = CDbl(DateSerial(y + 1, 1, 1))
            aIni = CDbl(DateSerial(y - 1, 1, 1)): aFim = pIni
            rot = "ano " & y & " (01/01 a 31/12)"
    End Select
    rotA = Format$(aIni, "dd/mm/yyyy") & " a " & Format$(aFim - 1, "dd/mm/yyyy")
    PN_Periodo = True
End Function

'----------------------------------------------------------
' Maquinas: cadastro (APS_Maquinas) + qualquer outra que apareca nas operacoes
'----------------------------------------------------------
Private Sub PN_ListarMaquinas()
    Dim m() As String, k As Long, i As Long, j As Long, ok As Boolean
    pnPasso = "PN_ListarMaquinas"
    k = APS_Maquinas(m)
    ReDim pnMaq(1 To k + IIf(pnN > 0, pnN, 1))
    pnNM = 0
    For i = 1 To k
        pnNM = pnNM + 1: pnMaq(pnNM) = m(i)
    Next i
    For i = 1 To pnN
        If Len(Trim$(pnOps(i).maquina)) > 0 Then
            ok = False
            For j = 1 To pnNM
                If APS_Ig(pnMaq(j), pnOps(i).maquina) Then ok = True: Exit For
            Next j
            If Not ok Then pnNM = pnNM + 1: pnMaq(pnNM) = Trim$(pnOps(i).maquina)
        End If
    Next i
End Sub

Private Function PN_IdxMaq(ByVal nome As String) As Long
    Dim j As Long
    pnPasso = "PN_IdxMaq"
    For j = 1 To pnNM
        If APS_Ig(pnMaq(j), nome) Then PN_IdxMaq = j: Exit Function
    Next j
End Function

' Marca as producoes envolvidas em conflito (inclusive quando o conflito e do bloco de Setup/Limpeza dela)
Private Sub PN_MarcarConflitos()
    Dim conf() As Boolean, i As Long, j As Long
    pnPasso = "PN_MarcarConflitos"
    ReDim pnConfP(1 To IIf(pnN > 0, pnN, 1))
    If pnN <= 0 Then Exit Sub
    APS_DetectarConflitos pnOps, pnN, conf
    For i = 1 To pnN
        If conf(i) Then
            If Len(pnOps(i).tipoAux) = 0 Then
                pnConfP(i) = True
            Else
                For j = 1 To pnN
                    If Len(pnOps(j).tipoAux) = 0 Then
                        If APS_Ig(pnOps(j).id, pnOps(i).idPai) Then pnConfP(j) = True: Exit For
                    End If
                Next j
            End If
        End If
    Next i
End Sub
'----------------------------------------------------------
' Parte de uma operacao dentro do periodo. Devolve False se nao ha sobreposicao.
'  wOv = tempo de TRABALHO (dias) dentro do periodo; fr = fracao da operacao (por tempo de trabalho) dentro do periodo
'----------------------------------------------------------
Private Function PN_Parte(ByRef o As tOperacao, ByVal pIni As Double, ByVal pFim As Double, _
                          ByRef wOv As Double, ByRef fr As Double) As Boolean
    Dim a As Double, b As Double, wTot As Double
    pnPasso = "PN_Parte"
    wOv = 0: fr = 0
    a = o.ini: If a < pIni Then a = pIni
    b = o.fim: If b > pFim Then b = pFim
    If b <= a + APS_EPS Then Exit Function
    PN_Parte = True
    wTot = APS_TrabalhoEntre(o.ini, o.fim)
    If o.ini >= pIni - APS_EPS And o.fim <= pFim + APS_EPS Then
        wOv = wTot: fr = 1
        Exit Function
    End If
    wOv = APS_TrabalhoEntre(a, b)
    If wTot > APS_EPS Then
        fr = wOv / wTot
    ElseIf o.fim - o.ini > APS_EPS Then
        fr = (b - a) / (o.fim - o.ini)
    End If
End Function

' Horas de trabalho do calendario no periodo (em blocos, porque APS_TrabalhoEntre limita 400 dias)
Private Function PN_Disp(ByVal pIni As Double, ByVal pFim As Double) As Double
    Dim a As Double, b As Double, t As Double
    pnPasso = "PN_Disp"
    a = pIni
    Do While a < pFim - APS_EPS
        b = a + 300
        If b > pFim Then b = pFim
        t = t + APS_TrabalhoEntre(a, b)
        a = b
    Loop
    PN_Disp = t * 24
End Function

Private Sub PN_SomaProd(ByRef a As tAgg, ByRef o As tOperacao, ByVal conf As Boolean, ByVal wOv As Double, ByVal fr As Double)
    pnPasso = "PN_SomaProd"
    If APS_Ig(o.status, APS_ST_CANCELADA) Then
        a.nCanc = a.nCanc + 1
        Exit Sub
    End If
    a.ordens = a.ordens + 1
    If APS_Ig(o.status, APS_ST_CONCLUIDA()) Then
        a.nConc = a.nConc + 1
    ElseIf APS_Ig(o.status, APS_ST_ANDAMENTO) Then
        a.nAnd = a.nAnd + 1
    ElseIf APS_Ig(o.status, APS_ST_ATRASADA) Then
        a.nAtr = a.nAtr + 1
    ElseIf Not APS_Ig(o.status, APS_ST_CONFLITO) Then
        a.nPlan = a.nPlan + 1
    End If
    If conf Or APS_Ig(o.status, APS_ST_CONFLITO) Then a.nConf = a.nConf + 1
    a.comp = a.comp + APS_CompDaOp(o) * fr
    a.hProd = a.hProd + wOv * 24
    If APS_Ig(o.status, APS_ST_ATRASADA) Or o.atraso > APS_EPS Then a.nAtrAny = a.nAtrAny + 1
    If APS_CompDaOp(o) <= APS_EPS Then a.nSemQtd = a.nSemQtd + 1
    If o.atraso > APS_EPS Then
        a.nAtrReg = a.nAtrReg + 1
        a.hAtrReg = a.hAtrReg + o.atraso * 24
    End If
End Sub

Private Sub PN_SomaAux(ByRef a As tAgg, ByRef o As tOperacao, ByVal wOv As Double)
    pnPasso = "PN_SomaAux"
    If APS_Ig(o.tipoAux, APS_TIPO_LIMPEZA) Then
        a.hLimp = a.hLimp + wOv * 24
        a.nLimp = a.nLimp + 1
    Else
        a.hSetup = a.hSetup + wOv * 24
        a.nSetup = a.nSetup + 1
    End If
End Sub

' Agrega o periodo: ag(0) = todas as maquinas; ag(k) = maquina k. Uma unica passada pelas operacoes.
Private Sub PN_Calc(ByVal pIni As Double, ByVal pFim As Double, ByRef ag() As tAgg)
    Dim i As Long, m As Long, wOv As Double, fr As Double
    pnPasso = "PN_Calc"
    ReDim ag(0 To pnNM)
    For i = 1 To pnN
        If PN_Parte(pnOps(i), pIni, pFim, wOv, fr) Then
            m = PN_IdxMaq(pnOps(i).maquina)
            If Len(pnOps(i).tipoAux) > 0 Then
                If APS_Ativa(pnOps(i)) Then
                    PN_SomaAux ag(0), pnOps(i), wOv
                    If m > 0 Then PN_SomaAux ag(m), pnOps(i), wOv
                End If
            Else
                PN_SomaProd ag(0), pnOps(i), pnConfP(i), wOv, fr
                If m > 0 Then PN_SomaProd ag(m), pnOps(i), pnConfP(i), wOv, fr
            End If
        End If
    Next i
End Sub
Private Function PN_HOcup(ByRef a As tAgg) As Double
    pnPasso = "PN_HOcup"
    PN_HOcup = a.hProd + a.hSetup + a.hLimp
End Function

' Ocupacao (fracao). -1 = nao calculavel (calendario sem horas de trabalho no periodo)
Private Function PN_Ocup(ByRef a As tAgg, ByVal disp As Double) As Double
    pnPasso = "PN_Ocup"
    If disp <= 0.0001 Then PN_Ocup = -1 Else PN_Ocup = PN_HOcup(a) / disp
End Function
Private Function PN_TxtOcup(ByVal o As Double) As String
    pnPasso = "PN_TxtOcup"
    If o < 0 Then PN_TxtOcup = "n/d" Else PN_TxtOcup = Format$(o, "0.0%")
End Function

Private Function PN_Causa(ByRef o As tOperacao) As String
    Dim s As String, p As Long
    pnPasso = "PN_Causa"
    s = o.obs
    p = InStr(s, "[[APS_")
    If p > 0 Then s = Left$(s, p - 1)
    s = Trim$(s)
    If Len(s) = 0 Then PN_Causa = "N" & ChrW(227) & "o registrada nos dados" Else PN_Causa = s
End Function

'----------------------------------------------------------
' Helpers de escrita (visual unico: titulos azuis, cinzas, vermelho so em alerta)
'----------------------------------------------------------
Private Sub PN_Limpar(ByVal ws As Worksheet)
    Dim i As Long
    pnPasso = "PN_Limpar"
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "pnl_" Then ws.Shapes(i).Delete
    Next i
    On Error Resume Next
    ws.Range("A" & PN_LIN0 & ":P1800").UnMerge
    ws.Range("A" & PN_LIN0 & ":P1800").Clear
    ws.Range(ws.Cells(PN_LIN0, PN_COLGRAF), ws.Cells(1800, 26)).Clear
    ws.Rows(PN_LIN0 & ":1800").RowHeight = ws.StandardHeight
    With ws.Range("A" & PN_LIN0 & ":P1800")
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CorTxt()
        .Interior.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
    pnFalhaGraf = ""
End Sub

' Titulo de secao: texto azul com linha fina embaixo
Private Sub PN_Titulo(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    pnPasso = "PN_Titulo"
    ws.Cells(r, 2).Value = txt
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, PN_COLN))
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = CorAz()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(191, 200, 210)
    End With
    ws.Rows(r).RowHeight = 26
End Sub

' Cabecalho de tabela: cinza, pequeno, sem preenchimento
Private Sub PN_Cab(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ParamArray t() As Variant)
    Dim i As Long
    pnPasso = "PN_Cab"
    For i = 0 To UBound(t)
        ws.Cells(r, c + i).Value = t(i)
    Next i
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + UBound(t)))
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = RGB(89, 89, 89)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(166, 177, 190)
    End With
    ws.Cells(r, c).HorizontalAlignment = xlLeft
    ws.Rows(r).RowHeight = 28
End Sub

Private Sub PN_Nota(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    pnPasso = "PN_Nota"
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Size = 9
    ws.Cells(r, 2).Font.Color = CorCinza()
End Sub

' Linha fina de separacao entre linhas de tabela
Private Sub PN_Fio(ByVal ws As Worksheet, ByVal r As Long)
    pnPasso = "PN_Fio"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, PN_COLN)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = CorLinha()
    End With
End Sub

' Botao (forma com macro). prim = True: azul com texto branco; False: cinza claro com texto azul.
Private Sub PN_Botao(ByVal ws As Worksheet, ByVal nome As String, ByVal txt As String, ByVal macro As String, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                     ByVal prim As Boolean, ByVal dica As String)
    Dim shp As Shape
    pnPasso = "PN_Botao"
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Name = nome
        .OnAction = macro
        .AlternativeText = dica
        .Placement = xlMove
        .Line.Visible = msoFalse
        If prim Then
            .Fill.ForeColor.RGB = CorAz()
        Else
            .Fill.ForeColor.RGB = RGB(226, 231, 237)
        End If
        With .TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = msoTrue
            If prim Then
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                .TextRange.Font.Fill.ForeColor.RGB = CorAz()
            End If
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 2
            .MarginRight = 2
        End With
    End With
End Sub

' Botao dentro de uma linha da tabela (ocupa as colunas c1 a c2 da linha r)
Private Sub PN_BotaoLinha(ByVal ws As Worksheet, ByVal nome As String, ByVal txt As String, ByVal macro As String, _
                          ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal dica As String, ByVal prim As Boolean)
    Dim x As Double, y As Double, w As Double, h As Double
    pnPasso = "PN_BotaoLinha"
    x = ws.Cells(r, c1).Left + 3
    y = ws.Cells(r, c1).Top + 3
    w = ws.Cells(r, c2 + 1).Left - x - 3
    h = ws.Cells(r, c1).Height - 6
    PN_Botao ws, nome, txt, macro, x, y, w, h, prim, dica
End Sub

' Cartao de indicador: rotulo, valor grande, ate duas linhas de explicacao. alerta = valor em vermelho.
Private Sub PN_Card(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal rotulo As String, _
                    ByVal valor As String, ByVal leg1 As String, ByVal leg2 As String, ByVal alerta As Boolean)
    Dim rg As Range, col As Range
    pnPasso = "PN_Card"
    Set rg = ws.Range(ws.Cells(r, c1), ws.Cells(r + 3, c2))
    Set col = ws.Range(ws.Cells(r, c1), ws.Cells(r + 3, c1))
    rg.Interior.Color = CorCartao()
    col.NumberFormat = "@"
    ws.Cells(r, c1).Value = rotulo
    ws.Cells(r + 1, c1).Value = valor
    ws.Cells(r + 2, c1).Value = leg1
    ws.Cells(r + 3, c1).Value = leg2
    col.HorizontalAlignment = xlLeft
    col.IndentLevel = 1
    On Error Resume Next
    ws.Cells(r + 1, c1).Errors(xlNumberAsText).Ignore = True
    ws.Cells(r, c1).Errors(xlNumberAsText).Ignore = True
    ws.Cells(r + 2, c1).Errors(xlNumberAsText).Ignore = True
    ws.Cells(r + 3, c1).Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    ws.Cells(r, c1).Font.Size = 10
    ws.Cells(r, c1).Font.Bold = True
    ws.Cells(r, c1).Font.Color = RGB(89, 89, 89)
    ws.Cells(r + 1, c1).Font.Size = 22
    ws.Cells(r + 1, c1).Font.Bold = True
    If alerta Then ws.Cells(r + 1, c1).Font.Color = CorAlerta() Else ws.Cells(r + 1, c1).Font.Color = CorTxt()
    ws.Range(ws.Cells(r + 2, c1), ws.Cells(r + 3, c1)).Font.Size = 8
    ws.Range(ws.Cells(r + 2, c1), ws.Cells(r + 3, c1)).Font.Color = CorCinza()
    ws.Rows(r).RowHeight = 22
    ws.Rows(r + 1).RowHeight = 34
    ws.Rows(r + 2).RowHeight = 14
    ws.Rows(r + 3).RowHeight = 14
    If c1 > 2 Then
        With rg.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 255, 255)
        End With
    End If
End Sub

' Grafico simples de uma serie (dados na area oculta). tipo: barra horizontal ou colunas.
Private Sub PN_Graf(ByVal ws As Worksheet, ByVal nome As String, ByVal titulo As String, ByVal tipo As Long, _
                    ByVal linTopo As Long, ByVal linAlt As Long, _
                    ByVal rX As Range, ByVal r1 As Range, ByVal n1 As String, ByVal fmtRot As String)
    Dim co As ChartObject, s As Series, L As Double, T As Double, W As Double, H As Double
    pnPasso = "PN_Graf"
    On Error GoTo Falha
    L = ws.Cells(linTopo, 2).Left
    T = ws.Cells(linTopo, 2).Top
    W = ws.Cells(linTopo, PN_COLN + 1).Left - L - 6
    H = ws.Cells(linTopo + linAlt, 2).Top - T - 6
    Set co = ws.ChartObjects.Add(L, T, W, H)
    co.Name = "pnl_" & nome
    With co.Chart
        .ChartType = tipo
        .PlotVisibleOnly = False
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = n1
        s.Values = r1
        s.XValues = rX
        s.Format.Fill.ForeColor.RGB = CorAz()
        If Len(fmtRot) > 0 Then
            s.HasDataLabels = True
            s.DataLabels.NumberFormat = fmtRot
        End If
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = titulo
        .ChartTitle.Font.Size = 11
    End With
    On Error Resume Next
    co.Chart.ChartTitle.Font.Bold = False
    co.Chart.ChartTitle.Font.Color = RGB(64, 64, 64)
    co.Chart.ChartArea.Format.Line.Visible = msoFalse
    co.Chart.Axes(xlValue).MinimumScale = 0
    co.Chart.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)
    If Len(fmtRot) > 0 Then co.Chart.Axes(xlValue).TickLabels.NumberFormat = fmtRot
    If tipo = xlBarClustered Then co.Chart.Axes(xlCategory).ReversePlotOrder = True
    On Error GoTo 0
    Exit Sub
Falha:
    pnFalhaGraf = pnFalhaGraf & titulo & " (" & Err.Description & "); "
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
End Sub

'----------------------------------------------------------
' Montagem: calcula e desenha a tela atual
'----------------------------------------------------------
Private Sub PN_Montar(ByVal ws As Worksheet)
    Dim msg As String, aviso As String, tela As String, modo As String, nm As String, r As Long
    Dim ok As Boolean

    pnPasso = "PN_Montar"
    APS_CalRecarregar
    pnN = APS_LerOps(pnOps)
    If pnN < 0 Then Err.Raise vbObjectError + 513, , "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel ler 02_Operacoes."
    If pnN = 0 Then ReDim pnOps(1 To 1)

    PN_ListarMaquinas

    tela = PN_Est(ws, 1)
    If StrComp(tela, PN_T_MAQ, vbTextCompare) <> 0 And StrComp(tela, PN_T_DET, vbTextCompare) <> 0 Then tela = PN_T_GERAL
    tela = LCase$(tela)
    modo = PN_Est(ws, 4)
    If StrComp(modo, PN_M_ATENCAO, vbTextCompare) <> 0 Then modo = PN_M_MAQ Else modo = PN_M_ATENCAO

    ' maquina em foco: na tela de detalhe vale a caixa de selecao; nas demais, a ultima consultada (so para destacar)
    nm = PN_Est(ws, 3)
    If StrComp(tela, PN_T_DET, vbTextCompare) = 0 And StrComp(modo, PN_M_MAQ, vbTextCompare) = 0 Then
        If Len(Trim$(CStr(ws.Range(PN_CMAQ).Value))) > 0 Then nm = Trim$(CStr(ws.Range(PN_CMAQ).Value))
    End If
    pnSel = 0
    If Len(nm) > 0 And StrComp(nm, "Todas", vbTextCompare) <> 0 Then pnSel = PN_IdxMaq(nm)
    If pnSel = 0 Then nm = "Todas" Else nm = pnMaq(pnSel)
    ws.Cells(3, PN_COLEST).Value = nm
    ws.Cells(1, PN_COLEST).Value = tela
    ws.Cells(4, PN_COLEST).Value = modo

    PN_MarcarConflitos
    pnNAl = APS_DiagnosticoColetar(pnAl)
    If pnNAl < 0 Then
        pnNAl = 0
        ReDim pnAl(1 To 1)
        msg = "O diagn" & ChrW(243) & "stico de dados n" & ChrW(227) & "o p" & ChrW(244) & "de ser executado."
    End If

    ok = PN_Periodo(ws, pnPIni, pnPFim, pnAIni, pnAFim, pnRot, pnRotA, aviso)
    PN_Limpar ws
    PN_Cabecalho ws, tela, modo, nm, aviso
    If Not ok Then Exit Sub
    If Len(msg) > 0 Then ws.Range("B6").Value = ws.Range("B6").Value & "  |  " & msg

    PN_Calc pnPIni, pnPFim, pnAgC
    PN_Calc pnAIni, pnAFim, pnAgP
    pnDispC = PN_Disp(pnPIni, pnPFim)
    pnDispP = PN_Disp(pnAIni, pnAFim)

    r = PN_LIN0
    If pnAgC(0).ordens = 0 Then
        PN_SemDados ws, r
        Exit Sub
    End If

    PN_ConstruirPri
    Select Case tela
        Case PN_T_GERAL
            PN_TelaGeral ws, r
        Case PN_T_MAQ
            PN_TelaMaquinas ws, r
        Case Else
            If StrComp(modo, PN_M_ATENCAO, vbTextCompare) = 0 Then
                PN_TelaAtencao ws, r
            Else
                PN_TelaMaquina ws, r
            End If
    End Select
    If Len(pnFalhaGraf) > 0 Then ws.Range("B6").Value = ws.Range("B6").Value & "  |  Gr" & ChrW(225) & "fico n" & ChrW(227) & "o gerado: " & pnFalhaGraf
End Sub

'----------------------------------------------------------
' Cabecalho fixo (linhas 1 a 7): titulo, onde estou, navegacao, filtros, periodo analisado
'----------------------------------------------------------
Private Sub PN_Cabecalho(ByVal ws As Worksheet, ByVal tela As String, ByVal modo As String, ByVal nm As String, ByVal aviso As String)
    Dim bc As String, ant As String, sep As String, x0 As Double, y0 As Double, i As Long
    Dim ehDet As Boolean, ehMaqDet As Boolean, ult As Long

    pnPasso = "PN_Cabecalho"
    ant = PN_Est(ws, 2)
    sep = " " & ChrW(8250) & " "
    ehDet = (StrComp(tela, PN_T_DET, vbTextCompare) = 0)
    ehMaqDet = (ehDet And StrComp(modo, PN_M_MAQ, vbTextCompare) = 0)

    ws.Rows(1).RowHeight = 32
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 28
    ws.Rows(4).RowHeight = 16
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 18
    ws.Rows(7).RowHeight = 8

    ws.Range("B1").Value = "Painel de produ" & ChrW(231) & "" & ChrW(227) & "o"
    ws.Range("B1").Font.Size = 20
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Color = CorAz()

    ' onde estou
    bc = "Vis" & ChrW(227) & "o geral"
    If StrComp(tela, PN_T_MAQ, vbTextCompare) = 0 Then
        bc = bc & sep & "M" & ChrW(225) & "quinas"
    ElseIf ehDet Then
        If ehMaqDet Then
            If StrComp(ant, PN_T_MAQ, vbTextCompare) = 0 Then bc = bc & sep & "M" & ChrW(225) & "quinas"
            If StrComp(nm, "Todas", vbTextCompare) = 0 Then
                bc = bc & sep & "Detalhes: todas as m" & ChrW(225) & "quinas"
            Else
                bc = bc & sep & "Detalhes: " & nm
            End If
        Else
            bc = bc & sep & "Detalhes: ordens que precisam de aten" & ChrW(231) & "" & ChrW(227) & "o"
        End If
    End If
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = bc
    ws.Range("B2").Font.Size = 11
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(64, 64, 64)

    ' navegacao (sempre visivel: Voltar so aparece fora da Visao geral)
    x0 = ws.Cells(3, 2).Left
    y0 = ws.Cells(3, 2).Top + 2
    If StrComp(tela, PN_T_GERAL, vbTextCompare) <> 0 Then
        PN_Botao ws, "pnl_btnVoltar", ChrW(8592) & " Voltar", "APS_PainelNavVoltar", x0, y0, 86, 24, False, "Volta para a tela anterior"
    End If
    PN_Botao ws, "pnl_btnGeral", "Vis" & ChrW(227) & "o geral", "APS_PainelNavGeral", x0 + 94, y0, 96, 24, (StrComp(tela, PN_T_GERAL, vbTextCompare) = 0), "Resumo do planejamento"
    PN_Botao ws, "pnl_btnMaquinas", "M" & ChrW(225) & "quinas", "APS_PainelNavMaquinas", x0 + 196, y0, 96, 24, (StrComp(tela, PN_T_MAQ, vbTextCompare) = 0), "Compara" & ChrW(231) & "" & ChrW(227) & "o entre as m" & ChrW(225) & "quinas"

    ' filtros (discretos)
    ws.Range("B4").Value = "Filtros"
    ws.Range("B5").Value = "Depois de alterar, clique em Atualizar."
    ws.Range("B4").Font.Bold = True
    ws.Range("B4:B5").Font.Size = 9
    ws.Range("B4:B5").Font.Color = CorCinza()
    ws.Range("C4").Value = "Per" & ChrW(237) & "odo"
    ws.Range("D4").Value = "Data de refer" & ChrW(234) & "ncia"
    ws.Range("C4:D4").Font.Size = 9
    ws.Range("C4:D4").Font.Bold = True
    ws.Range("C4:D4").Font.Color = CorCinza()
    With ws.Range("C5:D5")
        .Interior.Color = CorCartao()
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 200, 210)
        .HorizontalAlignment = xlCenter
        .Font.Color = CorTxt()
    End With
    ws.Range(PN_CREF).NumberFormat = "dd/mm/yyyy"
    With ws.Range(PN_CPER).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$R$1:$R$4"
    End With

    ws.Range("F4:G5").ClearContents
    ws.Range("F4:G5").Validation.Delete
    ws.Range("F4:G5").Interior.Color = RGB(255, 255, 255)
    ws.Range("F4:G5").Borders.LineStyle = xlNone
    If ehMaqDet Then
        ws.Range(ws.Cells(10, PN_COLLST), ws.Cells(80, PN_COLLST)).ClearContents
        ws.Cells(10, PN_COLLST).Value = "Todas"
        For i = 1 To pnNM
            ws.Cells(10 + i, PN_COLLST).Value = pnMaq(i)
        Next i
        ult = 10 + pnNM
        If ult > 80 Then ult = 80
        ws.Range("F4").Value = "M" & ChrW(225) & "quina"
        ws.Range("F4").Font.Size = 9
        ws.Range("F4").Font.Bold = True
        ws.Range("F4").Font.Color = CorCinza()
        With ws.Range(PN_CMAQ).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$R$10:$R$" & ult
        End With
        ws.Range(PN_CMAQ).Value = nm
        With ws.Range("F5:G5")
            .Interior.Color = CorCartao()
            .BorderAround xlContinuous, xlThin, , RGB(191, 200, 210)
        End With
        ws.Range(PN_CMAQ).HorizontalAlignment = xlLeft
        ws.Range(PN_CMAQ).IndentLevel = 1
    End If

    PN_Botao ws, "pnl_btnAtualizar", "Atualizar", "APS_PainelAtualizarBotao", ws.Cells(5, 9).Left, ws.Cells(5, 9).Top + 1, 90, ws.Cells(5, 9).Height - 2, True, "Recalcula o painel com os filtros escolhidos"

    ' periodo analisado (ou o erro, em vermelho)
    ws.Range("B6").NumberFormat = "@"
    ws.Range("B6").Font.Size = 9
    If Len(aviso) > 0 Then
        ws.Range("B6").Value = aviso
        ws.Range("B6").Font.Color = CorAlerta()
    Else
        ws.Range("B6").Value = "Per" & ChrW(237) & "odo analisado: " & pnRot & "   |   comparado com " & pnRotA & "   |   atualizado em " & Format$(Now, "dd/mm/yyyy hh:nn")
        ws.Range("B6").Font.Color = CorCinza()
    End If
End Sub

Private Sub PN_SemDados(ByVal ws As Worksheet, ByRef r As Long)
    Dim i As Long, minI As Double, maxF As Double, nOut As Long
    pnPasso = "PN_SemDados"
    For i = 1 To pnN
        If Len(pnOps(i).tipoAux) = 0 And APS_Ativa(pnOps(i)) Then
            If minI = 0 Or pnOps(i).ini < minI Then minI = pnOps(i).ini
            If pnOps(i).fim > maxF Then maxF = pnOps(i).fim
            nOut = nOut + 1
        End If
    Next i
    PN_Titulo ws, r, "Sem ordens no per" & ChrW(237) & "odo selecionado"
    r = r + 2
    ws.Cells(r, 2).Value = "N" & ChrW(227) & "o h" & ChrW(225) & " ordens planejadas (n" & ChrW(227) & "o canceladas) em " & pnRot & "."
    ws.Cells(r, 2).Font.Size = 11
    r = r + 1
    If pnAgC(0).nCanc > 0 Then
        ws.Cells(r, 2).Value = "Ordens canceladas no per" & ChrW(237) & "odo: " & pnAgC(0).nCanc & " (n" & ChrW(227) & "o entram em nenhum indicador)."
        r = r + 1
    End If
    If nOut > 0 Then
        ws.Cells(r, 2).Value = "H" & ChrW(225) & " " & nOut & " ordem(ns) em outros per" & ChrW(237) & "odos, de " & Format$(minI, "dd/mm/yyyy") & " at" & ChrW(233) & " " & Format$(maxF, "dd/mm/yyyy") & ". Ajuste o per" & ChrW(237) & "odo ou a data de refer" & ChrW(234) & "ncia e clique em Atualizar."
    Else
        ws.Cells(r, 2).Value = "N" & ChrW(227) & "o h" & ChrW(225) & " nenhuma ordem v" & ChrW(225) & "lida em 02_Operacoes."
    End If
End Sub

'----------------------------------------------------------
' Textos de situacao
'----------------------------------------------------------
Private Function PN_DeOrdens(ByVal n As Long) As String
    If n = 1 Then PN_DeOrdens = "de 1 ordem no per" & ChrW(237) & "odo" Else PN_DeOrdens = "de " & n & " ordens no per" & ChrW(237) & "odo"
End Function

Private Function PN_Junta(ByVal s As String, ByVal t As String) As String
    If Len(s) = 0 Then PN_Junta = t Else PN_Junta = s & " " & ChrW(183) & " " & t
End Function

' Situacao de uma maquina no periodo. Vazio = nada exige atencao.
Private Function PN_Situacao(ByRef a As tAgg, ByVal oc As Double) As String
    Dim s As String
    pnPasso = "PN_Situacao"
    If a.ordens = 0 Then Exit Function
    If a.nConf > 0 Then
        If a.nConf = 1 Then s = PN_Junta(s, "1 ordem em conflito") Else s = PN_Junta(s, a.nConf & " ordens em conflito")
    End If
    If a.nAtrAny > 0 Then
        If a.nAtrAny = 1 Then s = PN_Junta(s, "1 ordem atrasada") Else s = PN_Junta(s, a.nAtrAny & " ordens atrasadas")
    End If
    If oc > 1 + 0.000001 Then s = PN_Junta(s, "ocupa" & ChrW(231) & "" & ChrW(227) & "o acima de 100%")
    PN_Situacao = s
End Function

'----------------------------------------------------------
' Itens que precisam de atencao (lista unica, ordenada por prioridade)
'  1 conflito | 2 ordem atrasada | 3 dado inconsistente (alto) | 4 atraso de programacao | 5 maquina sobrecarregada | 6/7 dados (medio/baixo)
'----------------------------------------------------------
Private Sub PN_AddPri(ByRef L() As tPri, ByRef n As Long, ByVal pri As Long, ByVal tipo As String, ByVal tit As String, _
                      ByVal maq As String, ByVal ini As Double, ByVal fim As Double, ByVal txt As String)
    pnPasso = "PN_AddPri"
    n = n + 1
    If n > UBound(L) Then ReDim Preserve L(1 To UBound(L) + 100)
    L(n).pri = pri
    L(n).tipo = tipo
    L(n).tit = tit
    L(n).maq = maq
    L(n).ini = ini
    L(n).fim = fim
    L(n).txt = txt
End Sub

Private Function PN_NomeOp(ByRef o As tOperacao) As String
    pnPasso = "PN_NomeOp"
    Dim s As String
    If Len(Trim$(o.ordem)) > 0 Then s = "OP " & Trim$(o.ordem) Else s = "Sem OP"
    s = s & " " & ChrW(183) & " " & o.Produto
    If Len(Trim$(o.lote)) > 0 Then s = s & " (lote " & Trim$(o.lote) & ")"
    PN_NomeOp = s
End Function

Private Function PN_LinhaOp(ByRef o As tOperacao) As String
    pnPasso = "PN_LinhaOp"
    If o.linha > 0 Then PN_LinhaOp = " (linha " & o.linha & " em 02_Operacoes)"
End Function

Private Sub PN_ConstruirPri()
    Dim L() As tPri, nL As Long, i As Long, j As Long, k As Long, t As Long, ix() As Long
    Dim wOv As Double, fr As Double, lst As String, oc As Double, nome As String

    pnPasso = "PN_ConstruirPri"
    ReDim L(1 To pnN + pnNAl + pnNM + 10)
    For i = 1 To pnN
        If Len(pnOps(i).tipoAux) = 0 And APS_Ativa(pnOps(i)) Then
            If PN_Parte(pnOps(i), pnPIni, pnPFim, wOv, fr) Then
                nome = PN_NomeOp(pnOps(i))
                If pnConfP(i) Or APS_Ig(pnOps(i).status, APS_ST_CONFLITO) Then
                    lst = APS_ListaConflitos(pnOps, pnN, i)
                    If Len(lst) > 0 Then
                        lst = "Hor" & ChrW(225) & "rio sobreposto com: " & lst & "."
                    ElseIf pnConfP(i) Then
                        lst = "Conflito envolvendo o Setup/Limpeza desta ordem."
                    Else
                        lst = "Status marcado como Conflito."
                    End If
                    PN_AddPri L, nL, 1, "Conflito de hor" & ChrW(225) & "rio", nome, pnOps(i).maquina, pnOps(i).ini, pnOps(i).fim, _
                              lst & " Causa registrada: " & PN_Causa(pnOps(i)) & "." & PN_LinhaOp(pnOps(i))
                End If
                If APS_Ig(pnOps(i).status, APS_ST_ATRASADA) Then
                    PN_AddPri L, nL, 2, "Ordem atrasada", nome, pnOps(i).maquina, pnOps(i).ini, pnOps(i).fim, _
                              "Status Atrasada. Causa registrada: " & PN_Causa(pnOps(i)) & "." & PN_LinhaOp(pnOps(i))
                End If
                If pnOps(i).atraso > APS_EPS Then
                    PN_AddPri L, nL, 4, "Atraso de programa" & ChrW(231) & "" & ChrW(227) & "o", nome, pnOps(i).maquina, pnOps(i).ini, pnOps(i).fim, _
                              "Atraso de " & Format$(pnOps(i).atraso * 24, "0.0") & " h registrado na coluna Atraso. Causa registrada: " & PN_Causa(pnOps(i)) & "." & PN_LinhaOp(pnOps(i))
                End If
            End If
        End If
    Next i
    For k = 1 To pnNM
        If pnAgC(k).ordens > 0 Then
            oc = PN_Ocup(pnAgC(k), pnDispC)
            If oc > 1 + 0.000001 Then
                PN_AddPri L, nL, 5, "M" & ChrW(225) & "quina sobrecarregada", pnMaq(k), pnMaq(k), pnPIni, pnPFim - 1, _
                          Format$(PN_HOcup(pnAgC(k)), "0.0") & " h planejadas para " & Format$(pnDispC, "0.0") & " h de calend" & ChrW(225) & "rio (" & Format$(oc, "0.0%") & "). Veja os conflitos desta m" & ChrW(225) & "quina."
            End If
        End If
    Next k
    For i = 1 To pnNAl
        If Not APS_Ig(pnAl(i).cod, "SOBREP") Then
            If pnAl(i).ini = 0 Or (pnAl(i).fim > pnPIni And pnAl(i).ini < pnPFim) Then
                Select Case pnAl(i).sev
                    Case 1
                        t = 3
                        nome = "Dado inconsistente (alta)"
                    Case 2
                        t = 6
                        nome = "Dado inconsistente (m" & ChrW(233) & "dia)"
                    Case Else
                        t = 7
                        nome = "Dado inconsistente (baixa)"
                End Select
                If Len(pnAl(i).ref) > 0 Then lst = pnAl(i).ref Else lst = "Dados"
                PN_AddPri L, nL, t, nome, lst, "", pnAl(i).ini, pnAl(i).fim, pnAl(i).texto
            End If
        End If
    Next i

    ReDim ix(1 To IIf(nL > 0, nL, 1))
    For i = 1 To nL
        ix(i) = i
    Next i
    For i = 1 To nL - 1
        For j = i + 1 To nL
            If L(ix(j)).pri < L(ix(i)).pri Or (L(ix(j)).pri = L(ix(i)).pri And L(ix(j)).ini < L(ix(i)).ini) Then
                t = ix(i): ix(i) = ix(j): ix(j) = t
            End If
        Next j
    Next i
    ReDim pnPri(1 To IIf(nL > 0, nL, 1))
    pnNPri = nL
    For i = 1 To nL
        pnPri(i) = L(ix(i))
    Next i
End Sub

' Tabela de itens que precisam de atencao. completo = False: so os mais graves, sem detalhe (Visao geral).
Private Sub PN_TabPri(ByVal ws As Worksheet, ByRef r As Long, ByVal nMax As Long, ByVal completo As Boolean)
    Dim i As Long, n As Long, nEl As Long

    pnPasso = "PN_TabPri"
    If completo Then
        PN_Cab ws, r, 2, "Ordem ou item", "M" & ChrW(225) & "quina", "Problema", "", "In" & ChrW(237) & "cio", "O que est" & ChrW(225) & " acontecendo"
    Else
        PN_Cab ws, r, 2, "Ordem ou item", "M" & ChrW(225) & "quina", "Problema", "", "In" & ChrW(237) & "cio"
    End If
    ws.Range(ws.Cells(r, 3), ws.Cells(r, PN_COLN)).HorizontalAlignment = xlLeft
    r = r + 1
    For i = 1 To pnNPri
        If completo Or pnPri(i).pri <= 5 Then
            nEl = nEl + 1
            If n < nMax Then
                n = n + 1
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).NumberFormat = "@"
                ws.Cells(r, 2).Value = pnPri(i).tit
                On Error Resume Next
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Errors(xlNumberAsText).Ignore = True
                On Error GoTo 0
                ws.Cells(r, 3).Value = pnPri(i).maq
                ws.Cells(r, 4).Value = pnPri(i).tipo
                If pnPri(i).ini > 0 Then ws.Cells(r, 6).Value = pnPri(i).ini
                ws.Cells(r, 6).NumberFormat = "dd/mm/yy hh:mm"
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).HorizontalAlignment = xlLeft
                ws.Cells(r, 3).ShrinkToFit = True
                If pnPri(i).pri <= 2 Then ws.Cells(r, 4).Font.Color = CorAlerta()
                If completo Then
                    ws.Cells(r, 7).Value = pnPri(i).txt
                    ws.Range(ws.Cells(r, 7), ws.Cells(r, PN_COLN)).Merge
                    ws.Cells(r, 7).WrapText = True
                    ws.Cells(r, 7).Font.Size = 9
                    ws.Cells(r, 7).HorizontalAlignment = xlLeft
                    ws.Cells(r, 7).VerticalAlignment = xlTop
                    ws.Cells(r, 2).WrapText = True
                    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).VerticalAlignment = xlTop
                    ws.Rows(r).RowHeight = 40
                Else
                    ws.Cells(r, 2).ShrinkToFit = True
                    ws.Rows(r).RowHeight = 22
                End If
                PN_Fio ws, r
                r = r + 1
            End If
        End If
    Next i
    If nEl = 0 Then
        ws.Cells(r, 2).Value = "Nenhum conflito ou atraso neste per" & ChrW(237) & "odo."
        r = r + 1
    ElseIf nEl > n Then
        If completo Then
            PN_Nota ws, r, "... e mais " & (nEl - n) & " item(ns). Use o bot" & ChrW(227) & "o Diagn" & ChrW(243) & "stico de dados na Vis" & ChrW(227) & "o geral para a lista completa de dados."
        Else
            PN_Nota ws, r, "... e mais " & (nEl - n) & " item(ns). Clique em Ver todas."
        End If
        r = r + 1
    End If
End Sub

'----------------------------------------------------------
' TELA 1 - VISAO GERAL
'----------------------------------------------------------
Private Sub PN_TelaGeral(ByVal ws As Worksheet, ByRef r As Long)
    Dim a As tAgg, p As tAgg, oc As Double, op As Double, capC As Double, capP As Double
    Dim k As Long, n As Long, s As String, n1 As Long, i As Long, tx As String, rr As Long

    pnPasso = "PN_TelaGeral"
    a = pnAgC(0): p = pnAgP(0)
    capC = pnDispC * pnNM: capP = pnDispP * pnNM
    oc = PN_Ocup(a, capC): op = PN_Ocup(p, capP)

    ' --- como esta o planejamento
    PN_Titulo ws, r, "Como est" & ChrW(225) & " o planejamento"
    r = r + 1
    ws.Rows(r).RowHeight = 8
    r = r + 1
    If p.ordens > 0 Then tx = "Per" & ChrW(237) & "odo anterior: " & Format$(p.comp, "#,##0") Else tx = "Per" & ChrW(237) & "odo anterior: sem ordens"
    PN_Card ws, r, 2, 2, "Comprimidos planejados", Format$(a.comp, "#,##0"), "quantidade programada, n" & ChrW(227) & "o realizada", tx, False
    If p.ordens > 0 Then tx = "Per" & ChrW(237) & "odo anterior: " & PN_TxtOcup(op) Else tx = "Per" & ChrW(237) & "odo anterior: sem ordens"
    PN_Card ws, r, 3, 5, "Ocupa" & ChrW(231) & "" & ChrW(227) & "o das m" & ChrW(225) & "quinas", PN_TxtOcup(oc), "horas planejadas " & ChrW(247) & " horas do calend" & ChrW(225) & "rio", tx, (oc > 1 + 0.000001)
    PN_Card ws, r, 6, 8, "Ordens atrasadas", CStr(a.nAtrAny), "status Atrasada ou atraso registrado", PN_DeOrdens(a.ordens), (a.nAtrAny > 0)
    PN_Card ws, r, 9, 11, "Ordens em conflito", CStr(a.nConf), "hor" & ChrW(225) & "rios sobrepostos na mesma m" & ChrW(225) & "quina", PN_DeOrdens(a.ordens), (a.nConf > 0)
    r = r + 5

    ' --- maquinas que precisam de atencao
    PN_Titulo ws, r, "M" & ChrW(225) & "quinas que precisam de aten" & ChrW(231) & "" & ChrW(227) & "o"
    PN_BotaoLinha ws, "pnl_btnTodasMaq", "Ver todas as m" & ChrW(225) & "quinas", "APS_PainelNavMaquinas", r, 9, 11, "Abre a compara" & ChrW(231) & "" & ChrW(227) & "o entre m" & ChrW(225) & "quinas", False
    r = r + 1
    For k = 1 To pnNM
        s = PN_Situacao(pnAgC(k), PN_Ocup(pnAgC(k), pnDispC))
        If Len(s) > 0 Then
            n = n + 1
            If n <= 6 Then
                ws.Rows(r).RowHeight = 26
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 3).NumberFormat = "@"
                ws.Cells(r, 2).Value = pnMaq(k)
                ws.Cells(r, 2).Font.Bold = True
                ws.Cells(r, 3).Value = s
                ws.Cells(r, 3).Font.Color = CorAlerta()
                ws.Cells(r, 3).HorizontalAlignment = xlLeft
                PN_Fio ws, r
                PN_BotaoLinha ws, "pnl_maq_" & k, "Ver detalhes", "APS_PainelNavMaquina", r, 10, 11, pnMaq(k), False
                r = r + 1
            End If
        End If
    Next k
    If n = 0 Then
        ws.Cells(r, 2).Value = "Nenhuma m" & ChrW(225) & "quina precisa de aten" & ChrW(231) & "" & ChrW(227) & "o neste per" & ChrW(237) & "odo."
        r = r + 1
    ElseIf n > 6 Then
        PN_Nota ws, r, "... e mais " & (n - 6) & " m" & ChrW(225) & "quina(s). Clique em Ver todas as m" & ChrW(225) & "quinas."
        r = r + 1
    End If
    r = r + 1

    ' --- onde decidir
    PN_Titulo ws, r, "Ordens que exigem decis" & ChrW(227) & "o"
    PN_BotaoLinha ws, "pnl_btnTodasPri", "Ver todas", "APS_PainelNavAtencao", r, 10, 11, "Abre a lista completa com explica" & ChrW(231) & "" & ChrW(245) & "es", False
    r = r + 1
    PN_TabPri ws, r, 5, False
    r = r + 1

    ' --- rodape: o que os numeros significam
    PN_Nota ws, r, "Os valores s" & ChrW(227) & "o PLANEJADOS. O sistema n" & ChrW(227) & "o tem registro de produ" & ChrW(231) & "" & ChrW(227) & "o real; por isso produ" & ChrW(231) & "" & ChrW(227) & "o realizada, disponibilidade, desempenho e OEE n" & ChrW(227) & "o s" & ChrW(227) & "o exibidos."
    r = r + 1
    For i = 1 To pnNAl
        If pnAl(i).sev = 1 Then n1 = n1 + 1
    Next i
    PN_Nota ws, r, "Verifica" & ChrW(231) & "" & ChrW(227) & "o de dados: " & pnNAl & " alerta(s), " & n1 & " grave(s)."
    PN_BotaoLinha ws, "pnl_btnDiag", "Diagn" & ChrW(243) & "stico de dados", "APS_Diagnostico", r, 10, 11, "Abre o diagn" & ChrW(243) & "stico de dados (somente leitura)", False
    ws.Rows(r).RowHeight = 24
End Sub

'----------------------------------------------------------
' TELA 2 - MAQUINAS (comparacao e acesso ao detalhe)
'----------------------------------------------------------
Private Sub PN_TelaMaquinas(ByVal ws As Worksheet, ByRef r As Long)
    Dim k As Long, a As tAgg, oc As Double, r0 As Long, r1 As Long, ult As String, s As String, tot As tAgg, ocT As Double

    pnPasso = "PN_TelaMaquinas"
    ult = PN_Est(ws, 3)
    PN_Titulo ws, r, "Compara" & ChrW(231) & "" & ChrW(227) & "o entre m" & ChrW(225) & "quinas"
    r = r + 1
    PN_Nota ws, r, "Ocupa" & ChrW(231) & "" & ChrW(227) & "o = horas planejadas (produ" & ChrW(231) & "" & ChrW(227) & "o, setup e limpeza) " & ChrW(247) & " horas do calend" & ChrW(225) & "rio no per" & ChrW(237) & "odo. Horas livres = horas do calend" & ChrW(225) & "rio " & ChrW(8722) & " horas planejadas."
    r = r + 1
    PN_Cab ws, r, 2, "M" & ChrW(225) & "quina", "Situa" & ChrW(231) & "" & ChrW(227) & "o", "Ordens", "Comprimidos planejados", "Ocupa" & ChrW(231) & "" & ChrW(227) & "o", "Horas livres", "Ordens atrasadas", "Ordens em conflito"
    r = r + 1
    r0 = r
    ws.Range(ws.Cells(r0, PN_COLGRAF), ws.Cells(r0 + pnNM, PN_COLGRAF)).NumberFormat = "@"
    For k = 1 To pnNM
        a = pnAgC(k)
        oc = PN_Ocup(a, pnDispC)
        ws.Rows(r).RowHeight = 26
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = pnMaq(k)
        ws.Cells(r, 2).Font.Bold = True
        s = PN_Situacao(a, oc)
        If a.ordens = 0 Then
            ws.Cells(r, 3).Value = "Sem ordens"
            ws.Cells(r, 3).Font.Color = CorCinza()
        ElseIf Len(s) > 0 Then
            ws.Cells(r, 3).Value = "Aten" & ChrW(231) & "" & ChrW(227) & "o"
            ws.Cells(r, 3).Font.Color = CorAlerta()
            ws.Cells(r, 3).Font.Bold = True
        Else
            ws.Cells(r, 3).Value = "Normal"
        End If
        ws.Cells(r, 4).Value = a.ordens
        ws.Cells(r, 5).Value = a.comp
        If oc < 0 Then
            ws.Cells(r, 6).Value = "n/d"
            ws.Cells(r, 7).Value = "n/d"
        Else
            ws.Cells(r, 6).Value = oc
            If pnDispC - PN_HOcup(a) > 0 Then ws.Cells(r, 7).Value = pnDispC - PN_HOcup(a) Else ws.Cells(r, 7).Value = 0
            If oc > 1 + 0.000001 Then ws.Cells(r, 6).Font.Color = CorAlerta()
        End If
        ws.Cells(r, 8).Value = a.nAtrAny
        ws.Cells(r, 9).Value = a.nConf
        If a.nAtrAny > 0 Then ws.Cells(r, 8).Font.Color = CorAlerta()
        If a.nConf > 0 Then ws.Cells(r, 9).Font.Color = CorAlerta()
        ws.Cells(r, 4).NumberFormat = "0"
        ws.Cells(r, 5).NumberFormat = "#,##0"
        ws.Cells(r, 6).NumberFormat = "0.0%"
        ws.Cells(r, 7).NumberFormat = "0.0"
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).NumberFormat = "0"
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 9)).HorizontalAlignment = xlCenter
        If StrComp(pnMaq(k), ult, vbTextCompare) = 0 Then ws.Range(ws.Cells(r, 2), ws.Cells(r, PN_COLN)).Interior.Color = RGB(232, 240, 248)
        PN_Fio ws, r
        PN_BotaoLinha ws, "pnl_maq_" & k, "Ver detalhes", "APS_PainelNavMaquina", r, 10, 11, pnMaq(k), False
        ws.Cells(r, PN_COLGRAF).Value = pnMaq(k)
        If oc < 0 Then ws.Cells(r, PN_COLGRAF + 1).Value = 0 Else ws.Cells(r, PN_COLGRAF + 1).Value = oc
        r = r + 1
    Next k
    r1 = r - 1

    tot = pnAgC(0)
    ocT = PN_Ocup(tot, pnDispC * pnNM)
    ws.Rows(r).RowHeight = 24
    ws.Cells(r, 2).Value = "Todas as m" & ChrW(225) & "quinas"
    ws.Cells(r, 4).Value = tot.ordens
    ws.Cells(r, 5).Value = tot.comp
    If ocT < 0 Then ws.Cells(r, 6).Value = "n/d" Else ws.Cells(r, 6).Value = ocT
    ws.Cells(r, 8).Value = tot.nAtrAny
    ws.Cells(r, 9).Value = tot.nConf
    ws.Cells(r, 4).NumberFormat = "0"
    ws.Cells(r, 5).NumberFormat = "#,##0"
    ws.Cells(r, 6).NumberFormat = "0.0%"
    ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).NumberFormat = "0"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 9)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 9)).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, PN_COLN)).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r, PN_COLN)).Borders(xlEdgeTop).Color = RGB(166, 177, 190)
    r = r + 2

    If pnNM >= 2 Then
        pnPasso = "Maquinas: grafico"
        PN_Graf ws, "maqOcup", "Ocupa" & ChrW(231) & "" & ChrW(227) & "o planejada por m" & ChrW(225) & "quina", xlBarClustered, r, 12, _
                ws.Range(ws.Cells(r0, PN_COLGRAF), ws.Cells(r1, PN_COLGRAF)), ws.Range(ws.Cells(r0, PN_COLGRAF + 1), ws.Cells(r1, PN_COLGRAF + 1)), "Ocupa" & ChrW(231) & "" & ChrW(227) & "o", "0%"
        r = r + 13
    End If
End Sub

'----------------------------------------------------------
' TELA 3a - DETALHES DE UMA MAQUINA (ou de todas)
'----------------------------------------------------------
Private Sub PN_TelaMaquina(ByVal ws As Worksheet, ByRef r As Long)
    Dim a As tAgg, p As tAgg, oc As Double, op As Double, dC As Double, dP As Double, tit As String, tx As String
    Dim b As Long, r0 As Long, r1 As Long, cap As Double, v As Double
    Dim i As Long, j As Long, t As Long, n As Long, m As Long, ix() As Long, wOv As Double, fr As Double
    Dim vv() As Variant, k As Long, q As Long, setH As Double, flg As String, hSetLimp As Double, unid As String

    pnPasso = "PN_TelaMaquina"
    a = pnAgC(pnSel): p = pnAgP(pnSel)
    If pnSel > 0 Then
        tit = pnMaq(pnSel)
        dC = pnDispC: dP = pnDispP
    Else
        tit = "Todas as m" & ChrW(225) & "quinas"
        dC = pnDispC * pnNM: dP = pnDispP * pnNM
    End If
    oc = PN_Ocup(a, dC): op = PN_Ocup(p, dP)
    hSetLimp = a.hSetup + a.hLimp

    PN_Titulo ws, r, tit
    r = r + 1
    ws.Rows(r).RowHeight = 8
    r = r + 1
    If p.ordens > 0 Then tx = "Per" & ChrW(237) & "odo anterior: " & Format$(p.comp, "#,##0") Else tx = "Per" & ChrW(237) & "odo anterior: sem ordens"
    PN_Card ws, r, 2, 2, "Comprimidos planejados", Format$(a.comp, "#,##0"), "quantidade programada, n" & ChrW(227) & "o realizada", tx, False
    If p.ordens > 0 Then tx = "Per" & ChrW(237) & "odo anterior: " & PN_TxtOcup(op) Else tx = "Per" & ChrW(237) & "odo anterior: sem ordens"
    PN_Card ws, r, 3, 5, "Ocupa" & ChrW(231) & "" & ChrW(227) & "o", PN_TxtOcup(oc), "horas planejadas " & ChrW(247) & " horas do calend" & ChrW(225) & "rio", tx, (oc > 1 + 0.000001)
    PN_Card ws, r, 6, 8, "Setup e limpeza", Format$(hSetLimp, "0.0") & " h", a.nSetup & " setup(s) e " & a.nLimp & " limpeza(s)", "produ" & ChrW(231) & "" & ChrW(227) & "o: " & Format$(a.hProd, "0.0") & " h", False
    PN_Card ws, r, 9, 11, "Ordens no per" & ChrW(237) & "odo", CStr(a.ordens), a.nAtrAny & " atrasada(s)", a.nConf & " em conflito", False
    If a.nAtrAny > 0 Then ws.Cells(r + 2, 9).Font.Color = CorAlerta()
    If a.nConf > 0 Then ws.Cells(r + 3, 9).Font.Color = CorAlerta()
    r = r + 5

    ' --- evolucao (so quando ha pontos suficientes para ajudar)
    PN_CalcEvolucao
    If pnNB >= 3 Then
        pnPasso = "Detalhe: evolucao"
        If pnPFim - pnPIni <= 31 Then
            unid = "dia"
        ElseIf pnPFim - pnPIni <= 140 Then
            unid = "semana"
        Else
            unid = "m" & ChrW(234) & "s"
        End If
        PN_Titulo ws, r, "Ocupa" & ChrW(231) & "" & ChrW(227) & "o ao longo do per" & ChrW(237) & "odo"
        r = r + 1
        r0 = r
        ws.Range(ws.Cells(r0, PN_COLGRAF), ws.Cells(r0 + pnNB, PN_COLGRAF)).NumberFormat = "@"
        For b = 1 To pnNB
            If pnSel = 0 Then cap = pnBDisp(b) * pnNM Else cap = pnBDisp(b)
            If cap > 0.0001 Then v = pnBOc(b, pnSel) / cap Else v = 0
            ws.Cells(r, PN_COLGRAF).Value = pnBL(b)
            ws.Cells(r, PN_COLGRAF + 1).Value = v
            r = r + 1
        Next b
        r1 = r - 1
        r = r0
        PN_Graf ws, "detOcup", "Ocupa" & ChrW(231) & "" & ChrW(227) & "o planejada por " & unid, xlColumnClustered, r, 12, _
                ws.Range(ws.Cells(r0, PN_COLGRAF), ws.Cells(r1, PN_COLGRAF)), ws.Range(ws.Cells(r0, PN_COLGRAF + 1), ws.Cells(r1, PN_COLGRAF + 1)), "Ocupa" & ChrW(231) & "" & ChrW(227) & "o", "0%"
        r = r + 13
    End If

    ' --- ordens do periodo
    pnPasso = "Detalhe: ordens"
    ReDim ix(1 To IIf(pnN > 0, pnN, 1))
    For i = 1 To pnN
        If Len(pnOps(i).tipoAux) = 0 Then
            m = PN_IdxMaq(pnOps(i).maquina)
            If pnSel = 0 Or m = pnSel Then
                If PN_Parte(pnOps(i), pnPIni, pnPFim, wOv, fr) Then
                    n = n + 1
                    ix(n) = i
                End If
            End If
        End If
    Next i
    For i = 2 To n
        t = ix(i): j = i - 1
        Do While j >= 1
            If pnOps(ix(j)).ini <= pnOps(t).ini Then Exit Do
            ix(j + 1) = ix(j): j = j - 1
        Loop
        ix(j + 1) = t
    Next i

    PN_Titulo ws, r, "Ordens do per" & ChrW(237) & "odo (" & n & ")"
    r = r + 1
    If n = 0 Then
        ws.Cells(r, 2).Value = "Nenhuma ordem para esta m" & ChrW(225) & "quina no per" & ChrW(237) & "odo."
        r = r + 1
        Exit Sub
    End If
    PN_Cab ws, r, 2, "Produto", "OP", "Lote", "M" & ChrW(225) & "quina", "In" & ChrW(237) & "cio", "Fim", "Situa" & ChrW(231) & "" & ChrW(227) & "o", "Comprimidos planejados", "Setup e limpeza (h)", "Atraso (h)"
    r = r + 1

    q = n
    If q > PN_MAXDET Then q = PN_MAXDET
    ReDim vv(1 To q, 1 To 10)
    For k = 1 To q
        i = ix(k)
        Call PN_Parte(pnOps(i), pnPIni, pnPFim, wOv, fr)
        setH = 0
        For j = 1 To pnN
            If Len(pnOps(j).tipoAux) > 0 Then
                If APS_Ig(pnOps(j).idPai, pnOps(i).id) Then setH = setH + pnOps(j).dur * 24
            End If
        Next j
        flg = pnOps(i).status
        If pnConfP(i) Then flg = flg & " (conflito)"
        If Len(pnOps(i).codigo) > 0 Then vv(k, 1) = pnOps(i).codigo & " - " & pnOps(i).Produto Else vv(k, 1) = pnOps(i).Produto
        vv(k, 2) = pnOps(i).ordem
        vv(k, 3) = pnOps(i).lote
        vv(k, 4) = pnOps(i).maquina
        vv(k, 5) = pnOps(i).ini
        vv(k, 6) = pnOps(i).fim
        vv(k, 7) = flg
        If APS_Ig(pnOps(i).status, APS_ST_CANCELADA) Then vv(k, 8) = 0 Else vv(k, 8) = APS_CompDaOp(pnOps(i)) * fr
        vv(k, 9) = setH
        vv(k, 10) = pnOps(i).atraso * 24
    Next k
    ws.Range(ws.Cells(r, 2), ws.Cells(r + q - 1, 4)).NumberFormat = "@"
    ws.Range(ws.Cells(r, 2), ws.Cells(r + q - 1, 11)).Value = vv
    ws.Range(ws.Cells(r, 6), ws.Cells(r + q - 1, 7)).NumberFormat = "dd/mm/yy hh:mm"
    ws.Range(ws.Cells(r, 9), ws.Cells(r + q - 1, 9)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r, 10), ws.Cells(r + q - 1, 11)).NumberFormat = "0.0"
    ws.Range(ws.Cells(r, 3), ws.Cells(r + q - 1, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r, 2), ws.Cells(r + q - 1, 2)).ShrinkToFit = True
    ws.Range(ws.Cells(r, 5), ws.Cells(r + q - 1, 5)).ShrinkToFit = True
    ws.Range(ws.Cells(r, 8), ws.Cells(r + q - 1, 8)).ShrinkToFit = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r + q - 1, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r + q - 1, 11)).Borders(xlInsideHorizontal).Color = CorLinha()
    For k = 1 To q
        i = ix(k)
        If APS_Ig(pnOps(i).status, APS_ST_CANCELADA) Then
            ws.Range(ws.Cells(r + k - 1, 2), ws.Cells(r + k - 1, 11)).Font.Color = CorCinza()
        ElseIf pnConfP(i) Or APS_Ig(pnOps(i).status, APS_ST_ATRASADA) Or APS_Ig(pnOps(i).status, APS_ST_CONFLITO) Then
            ws.Cells(r + k - 1, 8).Font.Color = CorAlerta()
        End If
    Next k
    r = r + q
    If n > q Then
        PN_Nota ws, r, "... exibindo as primeiras " & q & " de " & n & " ordens. Reduza o per" & ChrW(237) & "odo ou escolha uma m" & ChrW(225) & "quina."
        r = r + 1
    End If
End Sub

'----------------------------------------------------------
' TELA 3b - DETALHES: ORDENS QUE PRECISAM DE ATENCAO
'----------------------------------------------------------
Private Sub PN_TelaAtencao(ByVal ws As Worksheet, ByRef r As Long)
    pnPasso = "PN_TelaAtencao"
    PN_Titulo ws, r, "Ordens e m" & ChrW(225) & "quinas que precisam de aten" & ChrW(231) & "" & ChrW(227) & "o"
    r = r + 1
    PN_Nota ws, r, "Em ordem de prioridade: conflitos de hor" & ChrW(225) & "rio, ordens atrasadas, dados inconsistentes graves, atrasos de programa" & ChrW(231) & "" & ChrW(227) & "o, m" & ChrW(225) & "quinas sobrecarregadas e demais avisos de dados."
    r = r + 1
    PN_Nota ws, r, "As causas s" & ChrW(227) & "o as registradas na coluna Observa" & ChrW(231) & "" & ChrW(245) & "es de 02_Operacoes; o painel n" & ChrW(227) & "o deduz causas."
    r = r + 2
    PN_TabPri ws, r, PN_MAXPRI, True
End Sub

'----------------------------------------------------------
' Evolucao por subperiodo (comprimidos e horas ocupadas por maquina)
'----------------------------------------------------------
Private Function PN_Buckets(ByVal pIni As Double, ByVal pFim As Double) As Long
    Dim n As Long, d As Double, e As Double, span As Double
    pnPasso = "PN_Buckets"
    ReDim pnBI(1 To 150): ReDim pnBF(1 To 150): ReDim pnBL(1 To 150)
    span = pFim - pIni
    d = pIni
    Do While d < pFim - APS_EPS And n < 150
        If span <= 31 Then
            e = d + 1
        ElseIf span <= 140 Then
            e = d + 7
        Else
            e = CDbl(DateSerial(Year(d), Month(d) + 1, 1))
        End If
        If e > pFim Then e = pFim
        n = n + 1
        pnBI(n) = d: pnBF(n) = e
        If span <= 140 Then pnBL(n) = Format$(d, "dd/mm") Else pnBL(n) = Format$(d, "mm/yyyy")
        d = e
    Loop
    PN_Buckets = n
End Function

Private Sub PN_CalcEvolucao()
    Dim b As Long, i As Long, m As Long, wOv As Double, fr As Double, cp As Double
    pnPasso = "PN_CalcEvolucao"
    pnNB = PN_Buckets(pnPIni, pnPFim)
    ReDim pnBComp(1 To pnNB, 0 To pnNM)
    ReDim pnBOc(1 To pnNB, 0 To pnNM)
    ReDim pnBDisp(1 To pnNB)
    For b = 1 To pnNB
        pnBDisp(b) = PN_Disp(pnBI(b), pnBF(b))
        For i = 1 To pnN
            If APS_Ativa(pnOps(i)) Then
                If PN_Parte(pnOps(i), pnBI(b), pnBF(b), wOv, fr) Then
                    m = PN_IdxMaq(pnOps(i).maquina)
                    If Len(pnOps(i).tipoAux) = 0 Then
                        cp = APS_CompDaOp(pnOps(i)) * fr
                        pnBComp(b, 0) = pnBComp(b, 0) + cp
                        If m > 0 Then pnBComp(b, m) = pnBComp(b, m) + cp
                    End If
                    pnBOc(b, 0) = pnBOc(b, 0) + wOv * 24
                    If m > 0 Then pnBOc(b, m) = pnBOc(b, m) + wOv * 24
                End If
            End If
        Next i
    Next b
End Sub
