Attribute VB_Name = "modAPS_VisaoMensal"
Option Explicit

'==========================================================
' APS - ABA "Visao Geral Mensal" (SOMENTE CONSULTA)
' Le a 02_Operacoes e mostra, para o mes/ano escolhido, o resumo por status e a tabela das ordens.
' NAO cria, apaga nem altera nada na 02_Operacoes (so le Value2 e escreve na propria aba).
' Entradas publicas:
'   APS_VM_Iniciar  : chamada por APS_AoAbrir (cria/atualiza a aba, mes e ano ATUAIS, liga os eventos)
'   APS_VM_Abrir    : macro manual (Alt+F8) / botao: faz o mesmo e ABRE a aba
'   APS_VM_Atualizar: recarrega com os filtros que estao na aba
'==========================================================

Private Const VM_L1 As Long = 11            ' primeira linha de dados da tabela
Private Const VM_CORTE As Long = 5000       ' limite de seguranca de linhas lidas por atualizacao

Private mEv As clsAPS_VM
Private mOcup As Boolean

Public Function APS_VM_NomeAba() As String
    APS_VM_NomeAba = "Vis" & ChrW(227) & "o Geral Mensal"
End Function

Private Function VM_Meses() As Variant
    VM_Meses = Array("Janeiro", "Fevereiro", "Mar" & ChrW(231) & "o", "Abril", "Maio", "Junho", _
                     "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
End Function

Private Function VM_Norm(ByVal s As String) As String
    Dim t As String, de As Variant, i As Long
    t = LCase$(Trim$(s))
    de = Array(225, 224, 226, 227, 233, 234, 237, 243, 244, 245, 250, 231)
    For i = 0 To 11
        t = Replace(t, ChrW(CLng(de(i))), Mid$("aaaaeeiooouc", i + 1, 1))
    Next i
    VM_Norm = t
End Function

Private Function VM_Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    VM_Txt = Trim$(CStr(v))
End Function

Private Function VM_Cel(ByRef dados As Variant, ByVal lin As Long, ByVal col As Long) As String
    If col < 1 Then Exit Function
    VM_Cel = VM_Txt(dados(lin, col))
End Function

Private Function VM_AchaCol(ByRef cab As Variant, ByVal n As Long, ByVal nome As String) As Long
    Dim j As Long, alvo As String
    alvo = VM_Norm(nome)
    For j = 1 To n
        If VM_Norm(VM_Txt(cab(1, j))) = alvo Then
            VM_AchaCol = j
            Exit Function
        End If
    Next j
End Function

' Linha de Setup/Limpeza gravada como se fosse ordem (modelo antigo): NAO e ordem de producao.
Private Function VM_EhSetup(ByVal cod As String, ByVal lote As String, ByVal obs As String) As Boolean
    If StrComp(cod, "SETUP", vbTextCompare) = 0 Then VM_EhSetup = True: Exit Function
    If StrComp(cod, "LIMPEZA", vbTextCompare) = 0 Then VM_EhSetup = True: Exit Function
    If InStr(1, obs, "[[APS_SETUP_DE", vbTextCompare) > 0 Then VM_EhSetup = True: Exit Function
    If Left$(UCase$(lote), 6) = "SETUP:" Then VM_EhSetup = True: Exit Function
End Function

Private Function VM_Aba() As Worksheet
    On Error Resume Next
    Set VM_Aba = ThisWorkbook.Worksheets(APS_VM_NomeAba())
    On Error GoTo 0
End Function

'==========================================================
' Estrutura da aba (idempotente: pode rodar de novo sem perder nada)
'==========================================================
Private Sub VM_Construir(ByVal ws As Worksheet)
    Dim cinza As Long, azul As Long, i As Long, rot As Variant, cor As Variant, ini As Variant

    azul = RGB(31, 58, 95)
    cinza = RGB(244, 246, 248)
    ws.Unprotect

    ws.Cells.Interior.Color = cinza
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 17
    ws.Columns("G").ColumnWidth = 17
    ws.Columns("H").ColumnWidth = 22
    ws.Columns("I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 3

    ' faixa de titulo
    With ws.Range("A1:J1")
        .Interior.Color = azul
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Rows(1).RowHeight = 30
    With ws.Range("B1")
        .Value2 = "VIS" & ChrW(195) & "O GERAL MENSAL"
        .Font.Size = 15
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("I1")
        .Value2 = "somente consulta"
        .Font.Size = 9
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ' botao CADASTROS E PARAMETROS (canto superior da aba, livre: nao fica sobre nenhum outro botao)
    Dim bp As Shape
    On Error Resume Next
    ws.Shapes("APS_BTN_PAR_VM").Delete
    On Error GoTo 0
    Set bp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Columns("F").Left + 4, ws.Rows(1).Top + 4, 170, 22)
    bp.Name = "APS_BTN_PAR_VM"
    bp.OnAction = "APS_AbrirParametros"
    bp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bp.Line.Visible = msoFalse
    bp.Placement = xlMove
    With bp.TextFrame2
        .TextRange.Text = "CADASTROS E PAR" & ChrW(194) & "METROS"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(31, 58, 95)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With ws.Range("B2")
        .Value2 = "Consulta do que est" & ChrW(225) & " na 02_Operacoes. N" & ChrW(227) & "o altera nenhum registro. Ordens que atravessam a virada do m" & ChrW(234) & "s aparecem nos dois meses."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(96, 125, 139)
    End With

    ' filtros
    rot = Array("M" & ChrW(234) & "s", "Ano", "M" & ChrW(225) & "quina (opcional)", "Status (opcional)")
    For i = 0 To 3
        With ws.Cells(3, 2 + i)
            .Value2 = rot(i)
            .Font.Size = 8
            .Font.Color = RGB(69, 90, 100)
            .VerticalAlignment = xlBottom
        End With
        With ws.Cells(4, 2 + i)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Locked = False
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(144, 164, 174)
        End With
    Next i
    ws.Rows(4).RowHeight = 24

    ' cartoes de resumo (celulas mescladas so para exibicao)
    rot = Array("Conclu" & ChrW(237) & "da", "Em Andamento", "Planejada", "Atrasada")
    cor = Array(RGB(46, 125, 50), RGB(21, 101, 192), RGB(96, 125, 139), RGB(198, 40, 40))
    ini = Array("B", "D", "F", "H")
    Dim fim As Variant
    fim = Array("C", "E", "G", "I")
    For i = 0 To 3
        ws.Range(ini(i) & "6:" & fim(i) & "6").Merge
        ws.Range(ini(i) & "7:" & fim(i) & "7").Merge
        With ws.Range(ini(i) & "6:" & fim(i) & "7")
            .Interior.Color = RGB(255, 255, 255)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeLeft).Color = cor(i)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(207, 216, 220)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(207, 216, 220)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(207, 216, 220)
        End With
        With ws.Range(ini(i) & "6")
            .Value2 = rot(i)
            .Font.Size = 9
            .Font.Color = RGB(69, 90, 100)
            .IndentLevel = 1
            .VerticalAlignment = xlCenter
        End With
        With ws.Range(ini(i) & "7")
            .NumberFormat = "0"
            .Font.Size = 22
            .Font.Bold = True
            .Font.Color = cor(i)
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Rows(6).RowHeight = 20
    ws.Rows(7).RowHeight = 34
    With ws.Range("B8")
        .Font.Size = 9
        .Font.Color = RGB(69, 90, 100)
    End With
    ws.Rows(8).RowHeight = 20

    With ws.Range("B9")
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(38, 50, 56)
    End With
    ws.Rows(9).RowHeight = 22

    ' cabecalho da tabela
    rot = Array("OP", "Produto", "Lote", "M" & ChrW(225) & "quina", "In" & ChrW(237) & "cio", "Fim", "Qtd. planejada (caixas)", "Status")
    For i = 0 To 7
        ws.Cells(10, 2 + i).Value2 = rot(i)
    Next i
    With ws.Range("B10:I10")
        .Interior.Color = azul
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("H10:I10").HorizontalAlignment = xlCenter
    ws.Rows(10).RowHeight = 30

    ' listas auxiliares (ocultas)
    ws.Columns("K:N").Hidden = True
    ws.Tab.Color = azul
End Sub

Private Sub VM_Validacao(ByVal c As Range, ByVal formula As String)
    With c.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formula
        .InCellDropdown = True
        .IgnoreBlank = False
        .ErrorTitle = "Vis" & ChrW(227) & "o Geral Mensal"
        .ErrorMessage = "Escolha um valor da lista."
    End With
End Sub

'==========================================================
' Cria a aba se nao existir (sem trocar a aba ativa do usuario)
'==========================================================
Private Function VM_Garantir() As Worksheet
    Dim ws As Worksheet, ant As Object

    Set ws = VM_Aba()
    If ws Is Nothing Then
        Set ant = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = APS_VM_NomeAba()
        If Not ant Is Nothing Then ant.Activate
    End If
    VM_Construir ws
    Set VM_Garantir = ws
End Function

Private Sub VM_LigarEventos()
    If mEv Is Nothing Then
        Set mEv = New clsAPS_VM
        Set mEv.App = Application
    End If
End Sub

'==========================================================
' Pontos de entrada
'==========================================================
Public Sub APS_VM_Iniciar()
    Dim ws As Worksheet
    On Error GoTo Falha
    Set ws = VM_Garantir()
    APS_VM_Definir Year(Date), Month(Date)
    VM_LigarEventos
    Exit Sub
Falha:
    MsgBox "Vis" & ChrW(227) & "o Geral Mensal: n" & ChrW(227) & "o foi poss" & ChrW(237) & "vel iniciar." & vbCrLf & Err.Description, vbExclamation, "APS"
End Sub

Public Sub APS_VM_Abrir()
    Dim ws As Worksheet
    On Error GoTo Falha
    Set ws = VM_Garantir()
    APS_VM_Definir Year(Date), Month(Date)
    VM_LigarEventos
    ws.Activate
    Exit Sub
Falha:
    MsgBox "Vis" & ChrW(227) & "o Geral Mensal: n" & ChrW(227) & "o foi poss" & ChrW(237) & "vel abrir." & vbCrLf & Err.Description, vbExclamation, "APS"
End Sub

' Define mes/ano (e zera os filtros opcionais) e recarrega.
Public Sub APS_VM_Definir(ByVal ano As Long, ByVal mes As Long)
    Dim ws As Worksheet, evt As Boolean
    Set ws = VM_Aba()
    If ws Is Nothing Then Exit Sub
    evt = Application.EnableEvents
    On Error GoTo Sai
    Application.EnableEvents = False
    ws.Unprotect
    ws.Range("B4").Value2 = VM_Meses()(mes - 1)
    ws.Range("C4").Value2 = ano
    ws.Range("D4").Value2 = "Todas"
    ws.Range("E4").Value2 = "Todos"
Sai:
    Application.EnableEvents = evt
    On Error GoTo 0
    APS_VM_Atualizar
End Sub

' Chamado quando a aba e ativada: recarrega mantendo os filtros escolhidos.
Public Sub APS_VM_AoAtivar()
    On Error Resume Next
    APS_VM_Atualizar
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

'==========================================================
' ATUALIZAR: le a 02_Operacoes, filtra e escreve a tabela
'==========================================================
Public Sub APS_VM_Atualizar()
    Dim ws As Worksheet, wsO As Worksheet
    Dim ano As Long, mes As Long, selMaq As String, selSt As String
    Dim mIni As Double, mFim As Double
    Dim cab As Variant, dados As Variant, nC As Long, nL As Long, ult As Long, i As Long, j As Long, k As Long
    Dim cCod As Long, cPro As Long, cOP As Long, cLot As Long, cMaq As Long, cIni As Long, cFim As Long
    Dim cSt As Long, cObs As Long, cCx As Long
    Dim ini As Double, fim As Double, st As String, nst As String, maq As String
    Dim keep() As Long, tmp As Long, out() As Variant
    Dim nConc As Long, nAnd As Long, nPla As Long, nAtr As Long, nOut As Long
    Dim totCx As Double, semCx As Long, semData As Long, ignSetup As Long
    Dim minAno As Long, maxAno As Long, lo As Long, hi As Long
    Dim maqs As Collection, listaMaq() As String, nMaq As Long, chave As String
    Dim evt As Boolean, passo As String, msg As String, vm As Variant, v As Variant, fixos As Variant
    Dim stSel As Variant

    If mOcup Then Exit Sub
    Set ws = VM_Aba()
    If ws Is Nothing Then Exit Sub
    mOcup = True
    evt = Application.EnableEvents
    On Error GoTo Falha

    passo = "preparar"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Unprotect

    ' ---- filtros ----
    passo = "ler filtros"
    vm = VM_Meses()
    mes = 0
    For i = 0 To 11
        If VM_Norm(VM_Txt(ws.Range("B4").Value2)) = VM_Norm(CStr(vm(i))) Then mes = i + 1
    Next i
    If mes = 0 Then mes = Month(Date)
    ano = 0
    If IsNumeric(ws.Range("C4").Value2) Then ano = CLng(Val(CStr(ws.Range("C4").Value2)))
    If ano < 1900 Or ano > 2200 Then ano = Year(Date)
    selMaq = VM_Txt(ws.Range("D4").Value2)
    If Len(selMaq) = 0 Or VM_Norm(selMaq) = "todas" Then selMaq = ""
    selSt = VM_Txt(ws.Range("E4").Value2)
    If Len(selSt) = 0 Or VM_Norm(selSt) = "todos" Then selSt = ""
    mIni = CDbl(DateSerial(ano, mes, 1))
    mFim = CDbl(DateSerial(ano, mes + 1, 1))

    ' ---- leitura da 02_Operacoes (somente leitura) ----
    passo = "abrir 02_Operacoes"
    On Error Resume Next
    Set wsO = ThisWorkbook.Worksheets("02_Operacoes")
    On Error GoTo Falha
    If wsO Is Nothing Then Err.Raise vbObjectError + 1, , "A aba 02_Operacoes n" & ChrW(227) & "o foi encontrada."

    passo = "cabecalho"
    nC = wsO.Cells(1, wsO.Columns.Count).End(xlToLeft).Column
    If nC < 2 Then Err.Raise vbObjectError + 2, , "Cabe" & ChrW(231) & "alho da 02_Operacoes n" & ChrW(227) & "o encontrado."
    cab = wsO.Range(wsO.Cells(1, 1), wsO.Cells(1, nC)).Value2
    cCod = VM_AchaCol(cab, nC, "C" & ChrW(243) & "digo do Produto")
    cPro = VM_AchaCol(cab, nC, "Produto")
    cOP = VM_AchaCol(cab, nC, "OP")
    cLot = VM_AchaCol(cab, nC, "Lote")
    cMaq = VM_AchaCol(cab, nC, "M" & ChrW(225) & "quina")
    cIni = VM_AchaCol(cab, nC, "Data de In" & ChrW(237) & "cio")
    cFim = VM_AchaCol(cab, nC, "Data de Fim")
    cSt = VM_AchaCol(cab, nC, "Planejamento")
    cObs = VM_AchaCol(cab, nC, "Observa" & ChrW(231) & ChrW(245) & "es")
    cCx = VM_AchaCol(cab, nC, "Quantidade de Caixas")
    msg = ""
    If cPro = 0 Then msg = msg & " Produto;"
    If cMaq = 0 Then msg = msg & " M" & ChrW(225) & "quina;"
    If cIni = 0 Then msg = msg & " Data de In" & ChrW(237) & "cio;"
    If cFim = 0 Then msg = msg & " Data de Fim;"
    If cSt = 0 Then msg = msg & " Planejamento;"
    If Len(msg) > 0 Then Err.Raise vbObjectError + 3, , "Coluna(s) n" & ChrW(227) & "o encontrada(s) na 02_Operacoes:" & msg

    passo = "ler dados"
    ult = wsO.Cells(wsO.Rows.Count, cIni).End(xlUp).Row
    If wsO.Cells(wsO.Rows.Count, cPro).End(xlUp).Row > ult Then ult = wsO.Cells(wsO.Rows.Count, cPro).End(xlUp).Row
    If ult > VM_CORTE Then ult = VM_CORTE
    nL = 0
    If ult >= 2 Then
        dados = wsO.Range(wsO.Cells(2, 1), wsO.Cells(ult, nC)).Value2
        nL = UBound(dados, 1)
    End If

    ' ---- filtragem ----
    passo = "filtrar"
    ReDim keep(1 To IIf(nL < 1, 1, nL))
    k = 0
    minAno = 0: maxAno = 0
    fixos = Array("Sala de Mistura", "FETTE 2090", "FETTE 2", "Mediseal", "Blisterflex", "Granula" & ChrW(231) & ChrW(227) & "o")
    Set maqs = New Collection
    nMaq = 0
    ReDim listaMaq(1 To 60)
    On Error Resume Next
    For j = 0 To UBound(fixos)
        Err.Clear
        maqs.Add 1, VM_Norm(CStr(fixos(j)))
        If Err.Number = 0 Then nMaq = nMaq + 1: listaMaq(nMaq) = CStr(fixos(j))
    Next j
    Err.Clear
    On Error GoTo Falha

    For i = 1 To nL
        If Len(VM_Txt(dados(i, cPro))) = 0 And Len(VM_Txt(dados(i, cMaq))) = 0 Then GoTo Prox
        If VM_EhSetup(VM_Cel(dados, i, cCod), VM_Cel(dados, i, cLot), VM_Cel(dados, i, cObs)) Then
            ignSetup = ignSetup + 1
            GoTo Prox
        End If
        If IsError(dados(i, cIni)) Or IsError(dados(i, cFim)) Then semData = semData + 1: GoTo Prox
        If Not IsNumeric(dados(i, cIni)) Or IsEmpty(dados(i, cIni)) Then semData = semData + 1: GoTo Prox
        If VarType(dados(i, cIni)) <> vbDouble Then semData = semData + 1: GoTo Prox
        ini = CDbl(dados(i, cIni))
        If VarType(dados(i, cFim)) = vbDouble Then fim = CDbl(dados(i, cFim)) Else fim = ini
        If fim < ini Then fim = ini

        ' anos e maquinas existentes na base (para as listas)
        If minAno = 0 Or Year(CDate(Int(ini))) < minAno Then minAno = Year(CDate(Int(ini)))
        If Year(CDate(Int(fim))) > maxAno Then maxAno = Year(CDate(Int(fim)))
        maq = VM_Txt(dados(i, cMaq))
        If Len(maq) > 0 Then
            chave = VM_Norm(maq)
            On Error Resume Next
            Err.Clear
            maqs.Add 1, chave
            If Err.Number = 0 Then
                If nMaq < 60 Then nMaq = nMaq + 1: listaMaq(nMaq) = maq
            End If
            Err.Clear
            On Error GoTo Falha
        End If

        ' atividade no mes (inicio antes do fim do mes E fim depois do inicio do mes; duracao zero conta pelo inicio)
        If ini < mFim Then
            If fim > mIni Or (fim = ini And ini >= mIni) Then
                st = VM_Txt(dados(i, cSt))
                maq = VM_Txt(dados(i, cMaq))
                If Len(selMaq) > 0 Then If VM_Norm(maq) <> VM_Norm(selMaq) Then GoTo Prox
                If Len(selSt) > 0 Then If VM_Norm(st) <> VM_Norm(selSt) Then GoTo Prox
                k = k + 1
                keep(k) = i
            End If
        End If
Prox:
    Next i

    ' ---- ordenacao: inicio, depois ordem original da planilha (estavel) ----
    passo = "ordenar"
    For i = 2 To k
        tmp = keep(i)
        j = i - 1
        Do While j >= 1
            If CDbl(dados(keep(j), cIni)) > CDbl(dados(tmp, cIni)) Then
                keep(j + 1) = keep(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        keep(j + 1) = tmp
    Next i

    ' ---- indicadores e tabela ----
    passo = "montar tabela"
    If k > 0 Then ReDim out(1 To k, 1 To 8)
    For i = 1 To k
        j = keep(i)
        st = VM_Txt(dados(j, cSt))
        nst = VM_Norm(st)
        If nst = "concluida" Then
            nConc = nConc + 1
        ElseIf nst = "em andamento" Then
            nAnd = nAnd + 1
        ElseIf nst = "planejada" Then
            nPla = nPla + 1
        ElseIf nst = "atrasada" Then
            nAtr = nAtr + 1
        Else
            nOut = nOut + 1
        End If
        out(i, 1) = VM_Cel(dados, j, cOP)
        If Len(out(i, 1)) = 0 Then out(i, 1) = ChrW(8212)
        out(i, 2) = VM_Txt(dados(j, cPro))
        out(i, 3) = VM_Cel(dados, j, cLot)
        out(i, 4) = VM_Txt(dados(j, cMaq))
        out(i, 5) = CDbl(dados(j, cIni))
        If VarType(dados(j, cFim)) = vbDouble Then out(i, 6) = CDbl(dados(j, cFim)) Else out(i, 6) = CDbl(dados(j, cIni))
        out(i, 7) = ChrW(8212)
        If cCx > 0 Then
            v = dados(j, cCx)
            If Not IsError(v) Then
                If VarType(v) = vbDouble Then
                    If v > 0 Then out(i, 7) = CDbl(v): totCx = totCx + CDbl(v)
                End If
            End If
        End If
        If VarType(out(i, 7)) <> vbDouble Then semCx = semCx + 1
        out(i, 8) = st
    Next i

    ' ---- escrita na aba (apenas B4:E4 normalizados + listas + resultado) ----
    passo = "escrever listas"
    lo = Year(Date) - 5: hi = Year(Date) + 5
    If minAno > 0 And minAno < lo Then lo = minAno
    If maxAno > hi Then hi = maxAno
    If ano < lo Then lo = ano
    If ano > hi Then hi = ano

    ws.Range("K1:N200").ClearContents
    ws.Range("K1").Value2 = "Meses": ws.Range("L1").Value2 = "Anos"
    ws.Range("M1").Value2 = "Maquinas": ws.Range("N1").Value2 = "Status"
    For i = 0 To 11
        ws.Cells(2 + i, 11).Value2 = vm(i)
    Next i
    For i = lo To hi
        ws.Cells(2 + i - lo, 12).Value2 = i
    Next i
    ws.Range("M2").Value2 = "Todas"
    For i = 1 To nMaq
        ws.Cells(2 + i, 13).Value2 = listaMaq(i)
    Next i
    stSel = Array("Todos", "Planejada", "Em Andamento", "Conclu" & ChrW(237) & "da", "Atrasada", "Cancelada", "Conflito")
    For i = 0 To 6
        ws.Cells(2 + i, 14).Value2 = stSel(i)
    Next i
    VM_Validacao ws.Range("B4"), "=$K$2:$K$13"
    VM_Validacao ws.Range("C4"), "=$L$2:$L$" & (2 + hi - lo)
    VM_Validacao ws.Range("D4"), "=$M$2:$M$" & (2 + nMaq)
    VM_Validacao ws.Range("E4"), "=$N$2:$N$8"

    passo = "escrever filtros"
    ws.Range("B4").Value2 = vm(mes - 1)
    ws.Range("C4").Value2 = ano
    If Len(selMaq) = 0 Then ws.Range("D4").Value2 = "Todas" Else ws.Range("D4").Value2 = selMaq
    If Len(selSt) = 0 Then ws.Range("E4").Value2 = "Todos" Else ws.Range("E4").Value2 = selSt

    passo = "escrever resumo"
    ws.Range("B7").Value2 = nConc
    ws.Range("D7").Value2 = nAnd
    ws.Range("F7").Value2 = nPla
    ws.Range("H7").Value2 = nAtr
    msg = "Total: " & k & IIf(k = 1, " ordem", " ordens") & "   |   Caixas planejadas (onde informada): " & Format$(totCx, "#,##0")
    If semCx > 0 Then msg = msg & " (" & semCx & " sem quantidade)"
    If nOut > 0 Then msg = msg & "   |   Outros status (Cancelada/Conflito/sem status): " & nOut
    If semData > 0 Then msg = msg & "   |   Ignoradas por data inv" & ChrW(225) & "lida: " & semData
    ws.Range("B8").Value2 = msg
    ws.Range("B9").Value2 = "Ordens de " & vm(mes - 1) & " de " & ano & " " & ChrW(8212) & " " & k & IIf(k = 1, " registro", " registros")

    passo = "escrever tabela"
    With ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + 3000, 9))
        .UnMerge
        .Clear
        .Interior.Color = RGB(244, 246, 248)
    End With
    ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + 3000, 9)).Font.Name = "Calibri"
    ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + 3000, 9)).Font.Size = 10

    If k = 0 Then
        With ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1, 9))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Font.Color = RGB(96, 125, 139)
            .Font.Italic = True
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
        ws.Rows(VM_L1).RowHeight = 30
        msg = "N" & ChrW(227) & "o existem ordens de produ" & ChrW(231) & ChrW(227) & "o para " & vm(mes - 1) & " de " & ano
        If Len(selMaq) > 0 Or Len(selSt) > 0 Then msg = msg & " com os filtros selecionados"
        ws.Cells(VM_L1, 2).Value2 = msg & "."
    Else
        ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + k - 1, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(VM_L1, 5), ws.Cells(VM_L1 + k - 1, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(VM_L1, 6), ws.Cells(VM_L1 + k - 1, 7)).NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Range(ws.Cells(VM_L1, 8), ws.Cells(VM_L1 + k - 1, 8)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(VM_L1, 9), ws.Cells(VM_L1 + k - 1, 9)).NumberFormat = "@"
        ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + k - 1, 9)).Value2 = out
        With ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + k - 1, 9))
            .Interior.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(236, 239, 241)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(207, 216, 220)
        End With
        ws.Range(ws.Cells(VM_L1, 6), ws.Cells(VM_L1 + k - 1, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(VM_L1, 8), ws.Cells(VM_L1 + k - 1, 8)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(VM_L1, 9), ws.Cells(VM_L1 + k - 1, 9)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(VM_L1, 2), ws.Cells(VM_L1 + k - 1, 2)).HorizontalAlignment = xlCenter
        ws.Rows(VM_L1 & ":" & (VM_L1 + k - 1)).RowHeight = 20
        For i = 1 To k
            nst = VM_Norm(CStr(out(i, 8)))
            With ws.Cells(VM_L1 + i - 1, 9)
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                Select Case nst
                    Case "concluida": .Interior.Color = RGB(46, 125, 50)
                    Case "em andamento": .Interior.Color = RGB(21, 101, 192)
                    Case "planejada": .Interior.Color = RGB(96, 125, 139)
                    Case "atrasada": .Interior.Color = RGB(198, 40, 40)
                    Case "conflito": .Interior.Color = RGB(239, 108, 0)
                    Case "cancelada": .Interior.Color = RGB(117, 117, 117)
                    Case Else: .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(96, 125, 139)
                End Select
            End With
        Next i
    End If

Saida:
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.EnableEvents = evt
    Application.ScreenUpdating = True
    mOcup = False
    Exit Sub

Falha:
    msg = "Vis" & ChrW(227) & "o Geral Mensal - erro na etapa '" & passo & "': " & Err.Description
    Err.Clear
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.EnableEvents = evt
    Application.ScreenUpdating = True
    mOcup = False
    MsgBox msg & vbCrLf & "Nenhum dado da 02_Operacoes foi alterado.", vbExclamation, "APS"
End Sub
