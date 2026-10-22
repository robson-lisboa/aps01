Attribute VB_Name = "modAPS_SAP"
Option Explicit

'==========================================================
' ABA SAP + ORGANIZAR AUTOMATICAMENTE
' A aba SAP e so uma area de entrada. Nada de regra nova:
'  - mes = 1_Planejamento!A1 (APS_LimiteMes)
'  - calendario e cascata = APS_Encaixar (motor existente)
'  - duracao = mesma regra do formulario (Mediseal: caixas + OEE; demais: duracao informada)
'  - ordem de processamento = data/hora pedida (ordenacao que o proprio motor ja usa); empate = ordem das linhas da aba SAP.
'    Nao e regra de prioridade de producao: a prioridade ainda sera informada pelo usuario.
' Nada e gravado sem confirmacao, e nada e gravado se alguma ordem nao couber no mes.
'==========================================================

Public Const APS_ABA_SAP As String = "SAP"

Private Const L_PRIM As Long = 5
Private Const L_ULT As Long = 204
Private Const C_ORD As Long = 1
Private Const C_COD As Long = 2
Private Const C_LOT As Long = 3
Private Const C_MAQ As Long = 4
Private Const C_DAT As Long = 5
Private Const C_HOR As Long = 6
Private Const C_DUR As Long = 7
Private Const C_QTD As Long = 8
Private Const C_OEE As Long = 9
Private Const C_OBS As Long = 10
Private Const C_SIT As Long = 11
Private Const C_PV As Long = 13

Private Type tSAP
    linha As Long
    ordem As String
    oper As tOperacao
    ok As Boolean
    erro As String
End Type

'----------------------------------------------------------
' Criacao da aba (so se ainda nao existir) e do botao
'----------------------------------------------------------
Public Sub APS_GarantirAbaSAP()
    Dim ws As Worksheet, ativa As Object

    On Error GoTo Sai
    Set ws = APS_Aba(APS_ABA_SAP)
    If ws Is Nothing Then
        Application.ScreenUpdating = False
        Set ativa = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = APS_ABA_SAP
        MontarAba ws
        If Not ativa Is Nothing Then ativa.Activate
    End If
    CriarBotaoSAP ws
Sai:
    Application.ScreenUpdating = True
End Sub

Private Sub CriarBotaoSAP(ByVal ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("APS_BTN_ORGANIZAR")
    On Error GoTo 0
    If shp Is Nothing Then
        Botao ws, "APS_BTN_ORGANIZAR", "ORGANIZAR AUTOMATICAMENTE", "APS_OrganizarAutomaticamente", _
              ws.Range("G1").Left, ws.Range("G1").Top + 3, 230, 24, RGB(46, 125, 50)
    End If
End Sub

Private Function MesTexto() As String
    Dim a As Double, b As Double
    If APS_LimiteMes(a, b) Then
        MesTexto = Format$(a, "mmmm/yyyy")
    Else
        MesTexto = "(n" & ChrW(227) & "o definido - preencha 1_Planejamento!A1)"
    End If
End Function

Private Sub MontarAba(ByVal ws As Worksheet)
    Dim h As Variant, i As Long, mq() As String, nm As Long, lista As String, larg As Variant

    ws.Range("A1").Value = "SAP - ORDENS PARA ORGANIZAR"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Rows(1).RowHeight = 30

    ws.Range("A2").Value = "Uma ordem por linha. M" & ChrW(225) & "quina, data, hora e dura" & ChrW(231) & "" & ChrW(227) & "o s" & ChrW(227) & "o obrigat" & ChrW(243) & "rias (na Mediseal: Qtd. Caixas e, se diferente do padr" & ChrW(227) & "o, OEE). O APS n" & ChrW(227) & "o completa dados nem cria regras."
    ws.Range("A2").Font.Italic = True
    ws.Range("A3").Value = "M" & ChrW(234) & "s do planejamento (1_Planejamento!A1): " & MesTexto()
    ws.Range("A3").Font.Bold = True

    h = Array("Ordem SAP", "C" & ChrW(243) & "digo do Produto", "Lote", "M" & ChrW(225) & "quina", "Data de In" & ChrW(237) & "cio", "Hora de In" & ChrW(237) & "cio", _
              "Dura" & ChrW(231) & "" & ChrW(227) & "o (hh:mm)", "Qtd. Caixas", "OEE (%)", "Observa" & ChrW(231) & "" & ChrW(245) & "es", "Situa" & ChrW(231) & "" & ChrW(227) & "o")
    larg = Array(14, 18, 14, 18, 15, 12, 15, 12, 9, 28, 44)
    For i = 0 To UBound(h)
        With ws.Cells(4, i + 1)
            .Value = h(i)
            .Interior.Color = APS_COR_CAB
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        ws.Columns(i + 1).ColumnWidth = larg(i)
    Next i

    ws.Range(ws.Cells(L_PRIM, C_DAT), ws.Cells(L_ULT, C_DAT)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(L_PRIM, C_HOR), ws.Cells(L_ULT, C_HOR)).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(L_PRIM, C_DUR), ws.Cells(L_ULT, C_DUR)).NumberFormat = "@"
    ws.Range(ws.Cells(L_PRIM, C_COD), ws.Cells(L_ULT, C_LOT)).NumberFormat = "@"
    ws.Range(ws.Cells(L_PRIM, C_ORD), ws.Cells(L_ULT, C_ORD)).NumberFormat = "@"

    nm = APS_Maquinas(mq)
    For i = 1 To nm
        If Len(lista) > 0 Then lista = lista & ","
        lista = lista & mq(i)
    Next i
    If Len(lista) > 0 Then
        With ws.Range(ws.Cells(L_PRIM, C_MAQ), ws.Cells(L_ULT, C_MAQ)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=lista
            .IgnoreBlank = True
        End With
    End If

    h = Array("Pr" & ChrW(233) & "via - opera" & ChrW(231) & "" & ChrW(227) & "o", "M" & ChrW(225) & "quina", "Antes", "Depois", "Motivo")
    larg = Array(40, 16, 30, 30, 44)
    For i = 0 To UBound(h)
        With ws.Cells(4, C_PV + i)
            .Value = h(i)
            .Interior.Color = RGB(84, 110, 122)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        ws.Columns(C_PV + i).ColumnWidth = larg(i)
    Next i
    ws.Columns(12).ColumnWidth = 3
End Sub

'----------------------------------------------------------
' Leitura e validacao das linhas da aba SAP (mesmas regras do formulario de operacao)
'----------------------------------------------------------
Private Function CelTexto(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        CelTexto = Format$(v, "hh:nn")
    ElseIf VarType(v) = vbDouble Then
        If v > 0 And v < 1 Then CelTexto = Format$(v, "hh:nn") Else CelTexto = CStr(v)
    Else
        CelTexto = Trim$(CStr(v))
    End If
End Function

Private Function ParaOEE(ByVal txt As String) As Double
    Dim s As String, v As Double
    s = Replace(Replace(Trim$(txt), "%", ""), ",", ".")
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9.]*" Then Exit Function
    v = Val(s)
    If v <= 0 Or v > 100 Then Exit Function
    ParaOEE = v / 100#
End Function

Private Function Validar(ByRef s As tSAP, ByVal v As Variant, ByVal r As Long, _
                         ByRef ops() As tOperacao, ByVal n As Long) As Boolean
    Dim cod As String, lote As String, maq As String, nome As String, outro As String
    Dim d As Double, h As Double, dur As Double, caixas As Double, oee As Double, velBase As Double
    Dim comp As Double, vEf As Double, mins As Double, i As Long, mq() As String, nm As Long, achou As Boolean
    Dim obs As String

    cod = CelTexto(v(r, C_COD))
    If Len(cod) = 0 Then s.erro = "Informe o c" & ChrW(243) & "digo do produto.": Exit Function
    nome = APS_NomeDoCodigo(cod)
    If Len(nome) = 0 Then s.erro = "C" & ChrW(243) & "digo de produto n" & ChrW(227) & "o cadastrado.": Exit Function

    lote = CelTexto(v(r, C_LOT))
    If Len(lote) = 0 Then s.erro = "Informe o lote.": Exit Function
    outro = APS_CodigoDoLote(ops, n, lote, "")
    If Len(outro) > 0 And outro <> cod Then
        s.erro = "O lote j" & ChrW(225) & " est" & ChrW(225) & " cadastrado para o produto " & outro & ".": Exit Function
    End If

    maq = CelTexto(v(r, C_MAQ))
    If Len(maq) = 0 Then s.erro = "Informe a m" & ChrW(225) & "quina.": Exit Function
    nm = APS_Maquinas(mq)
    For i = 1 To nm
        If APS_Ig(mq(i), maq) Then maq = mq(i): achou = True: Exit For
    Next i
    If Not achou Then s.erro = "M" & ChrW(225) & "quina n" & ChrW(227) & "o cadastrada.": Exit Function
    If Not APS_ProdutoPermitido(cod, maq) Then s.erro = "Este produto n" & ChrW(227) & "o pode ser programado nesta m" & ChrW(225) & "quina.": Exit Function

    d = APS_ParaData(v(r, C_DAT))
    If d < 0 Then s.erro = "Data inv" & ChrW(225) & "lida (use dd/mm/aaaa).": Exit Function
    h = APS_ParaHora(v(r, C_HOR))
    If h < 0 Then s.erro = "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido (use hh:mm).": Exit Function

    caixas = APS_ParaQtd(CelTexto(v(r, C_QTD)))
    If APS_EhMediseal(maq) Then
        velBase = APS_VelBaseDoCodigo(cod)
        If Len(CelTexto(v(r, C_OEE))) = 0 Then oee = APS_OEEPadrao() Else oee = ParaOEE(CelTexto(v(r, C_OEE)))
        If caixas <= 0 Then s.erro = "Informe a quantidade de caixas (Mediseal).": Exit Function
        If velBase <= 0 Then s.erro = "Velocidade-base n" & ChrW(227) & "o cadastrada para este produto.": Exit Function
        If oee <= 0 Then s.erro = "Informe um OEE v" & ChrW(225) & "lido (1 a 100%).": Exit Function
        If Not APS_CalcMediseal(caixas, velBase, oee, comp, vEf, mins) Then s.erro = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel calcular a dura" & ChrW(231) & "" & ChrW(227) & "o.": Exit Function
        dur = mins / 1440#
    Else
        dur = APS_ParaDuracao(CelTexto(v(r, C_DUR)))
        If dur <= 0 Then s.erro = "Informe a dura" & ChrW(231) & "" & ChrW(227) & "o (ex.: 10:30 ou 8h).": Exit Function
    End If

    ' ja existe em 02_Operacoes? (mesmo produto, lote e maquina)
    For i = 1 To n
        If APS_Ativa(ops(i)) Then
            If APS_Ig(ops(i).codigo, cod) And APS_Ig(ops(i).lote, lote) And APS_Ig(ops(i).maquina, maq) Then
                s.erro = "J" & ChrW(225) & " existe em 02_Operacoes (linha " & ops(i).linha & ").": Exit Function
            End If
        End If
    Next i

    ' Observacoes = so a observacao digitada; a Ordem SAP fica apenas na aba SAP
    obs = CelTexto(v(r, C_OBS))

    With s.oper
        .linha = 0
        .id = ""
        .codigo = cod
        .Produto = nome
        .lote = lote
        .maquina = maq
        .ini = APS_ArredMin(d + h)
        .dur = APS_ArredMin(dur)
        .Fim = APS_ArredMin(.ini + .dur)
        .caixas = caixas
        .status = APS_ST_PLANEJADA
        .obs = obs
        .oee = oee
        .velBase = velBase
        .SemDur = False
    End With
    Validar = True
End Function

Private Function Faixa(ByVal ini As Double, ByVal Fim As Double) As String
    Faixa = Format$(ini, "dd/mm hh:nn") & " - " & Format$(Fim, "dd/mm hh:nn")
End Function

Private Function rotulo(ByRef o As tOperacao, ByVal ordem As String) As String
    If Len(ordem) > 0 Then rotulo = "OP SAP " & ordem & " - " Else rotulo = ""
    rotulo = rotulo & o.Produto & " " & o.lote
End Function

'----------------------------------------------------------
' ORGANIZAR AUTOMATICAMENTE
'----------------------------------------------------------
Public Sub APS_OrganizarAutomaticamente()
    Dim wsS As Worksheet, wsO As Worksheet, est As Boolean, trav As Boolean
    Dim mIni As Double, mFim As Double
    Dim ops0() As tOperacao, n0 As Long
    Dim ops() As tOperacao, orig() As tOperacao, mot() As String, pedido() As Double, rot() As String
    Dim s() As tSAP, ns As Long, ord() As Long, v As Variant
    Dim i As Long, j As Long, k As Long, t As Long, nc As Long, r As Long, x As Long
    Dim mudou() As Boolean, calOk As Boolean, qtd As Long, conf() As Boolean
    Dim nErr As Long, msg As String, pv() As String, np As Long, nNova As Long, nMov As Long, nConf As Long
    Dim resumo As String, lista As String, res As VbMsgBoxResult, motivo As String

    On Error GoTo Falha

    Set wsS = APS_Aba(APS_ABA_SAP)
    If wsS Is Nothing Then
        MsgBox "A aba SAP n" & ChrW(227) & "o existe.", vbExclamation, APS_TITULO
        Exit Sub
    End If
    If APS_Aba(APS_ABA_OPS) Is Nothing Or APS_Aba(APS_ABA_PLAN) Is Nothing Then Exit Sub

    ' 7) mes de planejamento obrigatorio
    If Not APS_LimiteMes(mIni, mFim) Then
        MsgBox "Defina o M" & ChrW(234) & "s do Planejamento em 1_Planejamento, c" & ChrW(233) & "lula A1 (ex.: 01/09/2026), antes de organizar.", _
               vbExclamation, APS_TITULO
        Exit Sub
    End If
    wsS.Range("A3").Value = "M" & ChrW(234) & "s do planejamento (1_Planejamento!A1): " & MesTexto()

    ' limpa a previa anterior
    wsS.Range(wsS.Cells(5, C_PV), wsS.Cells(L_ULT + 200, C_PV + 4)).ClearContents

    ' 3) planejamento atual
    APS_CalRecarregar
    APS_GarantirIDs
    n0 = APS_LerOps(ops0)
    If n0 < 0 Then Exit Sub

    ' 1, 2, 5) ordens da aba SAP
    v = wsS.Range(wsS.Cells(L_PRIM, 1), wsS.Cells(L_ULT, C_SIT)).Value
    ReDim s(1 To L_ULT - L_PRIM + 1)
    For r = 1 To UBound(v, 1)
        If Left$(CelTexto(v(r, C_SIT)), 8) <> "Aplicada" Then
            If Len(CelTexto(v(r, C_ORD))) + Len(CelTexto(v(r, C_COD))) + Len(CelTexto(v(r, C_LOT))) + _
               Len(CelTexto(v(r, C_MAQ))) + Len(CelTexto(v(r, C_QTD))) + Len(CelTexto(v(r, C_DUR))) > 0 Then
                ns = ns + 1
                s(ns).linha = L_PRIM + r - 1
                s(ns).ordem = CelTexto(v(r, C_ORD))
                s(ns).ok = Validar(s(ns), v, r, ops0, n0)
                If Not s(ns).ok Then nErr = nErr + 1
            End If
        End If
    Next r

    If ns = 0 Then
        MsgBox "Nenhuma ordem pendente na aba SAP.", vbInformation, APS_TITULO
        Exit Sub
    End If

    ' mesmo lote com produtos diferentes dentro da propria aba
    For i = 1 To ns
        If s(i).ok Then
            For j = 1 To i - 1
                If s(j).ok Then
                    If APS_Ig(s(i).oper.lote, s(j).oper.lote) And Not APS_Ig(s(i).oper.codigo, s(j).oper.codigo) Then
                        s(i).ok = False
                        s(i).erro = "O lote " & s(i).oper.lote & " aparece com outro produto na linha " & s(j).linha & "."
                        nErr = nErr + 1
                        Exit For
                    End If
                    If APS_Ig(s(i).oper.lote, s(j).oper.lote) And APS_Ig(s(i).oper.codigo, s(j).oper.codigo) And _
                       APS_Ig(s(i).oper.maquina, s(j).oper.maquina) Then
                        s(i).ok = False
                        s(i).erro = "Duplicada na linha " & s(j).linha & "."
                        nErr = nErr + 1
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    If nErr > 0 Then
        For i = 1 To ns
            If s(i).ok Then
                wsS.Cells(s(i).linha, C_SIT).Value = "OK (nada gravado: corrija as demais linhas)"
            Else
                wsS.Cells(s(i).linha, C_SIT).Value = "Erro: " & s(i).erro
            End If
        Next i
        MsgBox nErr & " linha(s) com problema na aba SAP (coluna Situa" & ChrW(231) & "" & ChrW(227) & "o). Nada foi alterado." & vbCrLf & _
               "Informe/corrija os dados e clique novamente.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    ' ordem de processamento (nao e prioridade): data/hora pedida, empate = ordem das linhas da aba SAP
    ReDim ord(1 To ns)
    For i = 1 To ns: ord(i) = i: Next i
    For i = 2 To ns
        t = ord(i): j = i - 1
        Do While j >= 1
            If s(ord(j)).oper.ini > s(t).oper.ini + APS_EPS Then
                ord(j + 1) = ord(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        ord(j + 1) = t
    Next i

    ' 4, 6, 8, 9, 10, 11) simulacao em memoria com o motor existente (nada e gravado aqui)
    ReDim ops(1 To n0 + ns)
    ReDim orig(1 To n0 + ns)
    ReDim mot(1 To n0 + ns)
    ReDim pedido(1 To n0 + ns)
    ReDim rot(1 To n0 + ns)
    For i = 1 To n0
        ops(i) = ops0(i)
        orig(i) = ops0(i)
    Next i

    nc = n0
    For t = 1 To ns
        nc = nc + 1
        k = nc
        ops(k) = s(ord(t)).oper
        orig(k) = ops(k)
        pedido(k) = ops(k).ini
        rot(k) = rotulo(ops(k), s(ord(t)).ordem)
        APS_Encaixar ops, nc, k, mudou, calOk
        If Not calOk Then
            If APS_ForaDoMes Then msg = APS_MsgForaDoMes() Else msg = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui per" & ChrW(237) & "odo de trabalho para encaixar a opera" & ChrW(231) & "" & ChrW(227) & "o."
            For i = 1 To ns
                wsS.Cells(s(i).linha, C_SIT).Value = IIf(i = ord(t), "N" & ChrW(227) & "o coube: " & IIf(APS_ForaDoMes, "fora do m" & ChrW(234) & "s do planejamento", "calend" & ChrW(225) & "rio"), "Nada gravado")
            Next i
            MsgBox "Ordem " & rot(k) & ":" & vbCrLf & msg & vbCrLf & vbCrLf & "Nenhuma organiza" & ChrW(231) & "" & ChrW(227) & "o foi gravada.", vbExclamation, APS_TITULO
            Exit Sub
        End If
        For j = 1 To nc - 1
            If mudou(j) Then
                mot(j) = "cascata: empurrada por " & rot(k)
            End If
        Next j
    Next t

    ' 4) conflitos restantes (ex.: operacao que nao pode ser movida)
    APS_DetectarConflitos ops, nc, conf

    ' previa (colunas M:Q da aba SAP)
    ReDim pv(1 To ns + n0, 1 To 5)
    np = 0
    For t = 1 To ns
        k = n0 + t
        np = np + 1: nNova = nNova + 1
        pv(np, 1) = "NOVA  " & rot(k)
        pv(np, 2) = ops(k).maquina
        pv(np, 3) = "pedido " & Faixa(pedido(k), pedido(k) + ops(k).dur)
        pv(np, 4) = Faixa(ops(k).ini, ops(k).Fim)
        If Abs(ops(k).ini - pedido(k)) <= APS_EPS Then
            motivo = "sem conflito"
        ElseIf APS_ProximoDisponivel(pedido(k)) > pedido(k) + APS_EPS Then
            motivo = "calend" & ChrW(225) & "rio (fora do per" & ChrW(237) & "odo de trabalho) e/ou conflito de hor" & ChrW(225) & "rio"
        Else
            motivo = "conflito de hor" & ChrW(225) & "rio na m" & ChrW(225) & "quina"
        End If
        If Len(mot(k)) > 0 Then motivo = motivo & "; " & mot(k)
        lista = APS_ListaConflitos(ops, nc, k)
        If Len(lista) > 0 Then motivo = motivo & " | CONFLITO restante com: " & lista: nConf = nConf + 1
        pv(np, 5) = motivo
    Next t
    For j = 1 To n0
        If Abs(ops(j).ini - orig(j).ini) > APS_EPS Then
            np = np + 1: nMov = nMov + 1
            pv(np, 1) = "Existente  " & ops(j).Produto & " " & ops(j).lote & " (linha " & ops(j).linha & ")"
            pv(np, 2) = ops(j).maquina
            pv(np, 3) = Faixa(orig(j).ini, orig(j).Fim)
            pv(np, 4) = Faixa(ops(j).ini, ops(j).Fim)
            pv(np, 5) = IIf(Len(mot(j)) > 0, mot(j), "conflito de hor" & ChrW(225) & "rio")
        End If
    Next j
    wsS.Range(wsS.Cells(5, C_PV), wsS.Cells(4 + np, C_PV + 4)).Value = SubMatriz(pv, np)

    ' resumo simples (formato: antes / depois / motivo)
    x = 0
    For i = 1 To np
        If x < 3 Then
            resumo = resumo & pv(i, 1) & vbCrLf & "  Antes: " & pv(i, 3) & vbCrLf & "  Depois: " & pv(i, 4) & vbCrLf & "  Motivo: " & Left$(pv(i, 5), 70) & vbCrLf & vbCrLf
            x = x + 1
        End If
    Next i
    If np > x Then resumo = resumo & "... e mais " & (np - x) & " linha(s) na pr" & ChrW(233) & "via (aba SAP, colunas M a Q)." & vbCrLf & vbCrLf

    res = MsgBox("PR" & ChrW(201) & "VIA DA ORGANIZA" & ChrW(199) & "" & ChrW(195) & "O - m" & ChrW(234) & "s " & Format$(mIni, "mmmm/yyyy") & vbCrLf & vbCrLf & _
                 nNova & " opera" & ChrW(231) & "" & ChrW(227) & "o(" & ChrW(245) & "es) nova(s), " & nMov & " existente(s) deslocada(s)." & vbCrLf & _
                 IIf(nConf > 0, "ATEN" & ChrW(199) & "" & ChrW(195) & "O: " & nConf & " com conflito restante (opera" & ChrW(231) & "" & ChrW(227) & "o que n" & ChrW(227) & "o pode ser movida)." & vbCrLf, "") & vbCrLf & _
                 resumo & "SIM = APLICAR ORGANIZA" & ChrW(199) & "" & ChrW(195) & "O" & vbCrLf & "N" & ChrW(195) & "O = CANCELAR (nada " & ChrW(233) & " gravado)", _
                 vbYesNo + vbQuestion + vbDefaultButton2, APS_TITULO)

    If res <> vbYes Then
        For i = 1 To ns
            wsS.Cells(s(i).linha, C_SIT).Value = "Pr" & ChrW(233) & "via n" & ChrW(227) & "o aplicada"
        Next i
        Exit Sub
    End If

    ' 12) aplica: 02_Operacoes e depois cards
    Set wsO = APS_Aba(APS_ABA_OPS)
    est = APS_Liberar(wsO)
    APS_Ocupado = APS_Ocupado + 1
    trav = True
    For t = 1 To ns
        k = n0 + t
        APS_GravarOp ops(k)
        wsS.Cells(s(ord(t)).linha, C_SIT).Value = "Aplicada " & Format$(Now, "dd/mm/yyyy hh:nn") & " -> OP " & ops(k).id
    Next t
    For j = 1 To n0
        If Abs(ops(j).ini - orig(j).ini) > APS_EPS Then APS_GravarHorario ops(j)
    Next j
    APS_Ocupado = APS_Ocupado - 1
    trav = False
    APS_Reproteger wsO, est
    est = False

    APS_Atualizar True
    MsgBox "Organiza" & ChrW(231) & "" & ChrW(227) & "o aplicada: " & nNova & " nova(s), " & nMov & " existente(s) deslocada(s).", vbInformation, APS_TITULO
    Exit Sub

Falha:
    msg = Err.Description
    On Error Resume Next
    If trav Then APS_Ocupado = APS_Ocupado - 1
    If est Then APS_Reproteger wsO, est
    MsgBox "Erro ao organizar: " & msg, vbCritical, APS_TITULO
End Sub

' primeiras np linhas da matriz de previa
Private Function SubMatriz(ByRef pv() As String, ByVal np As Long) As Variant
    Dim m() As String, i As Long, c As Long
    If np < 1 Then np = 1
    ReDim m(1 To np, 1 To 5)
    For i = 1 To np
        For c = 1 To 5
            m(i, c) = pv(i, c)
        Next c
    Next i
    SubMatriz = m
End Function




