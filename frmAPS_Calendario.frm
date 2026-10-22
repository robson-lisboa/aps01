VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPS_Calendario 
   Caption         =   "Calendario"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAPS_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' APS PURAN - CALENDARIO (versao compacta)
'  Esquerda : semana padrao (cada dia independente) e OEE padrao.
'  Direita  : bloqueios. 1) clique numa data (ou em duas, para um periodo)  2) escolha "Dia inteiro" ou "Horario"  3) Bloquear.
'
' Grava nos MESMOS nomes que o motor ja le (nada mudou no motor nem no calendario):
'   CAL_SEM  semana padrao            CAL_EXC  excecoes por data (dia inteiro = "dd/mm/aaaa;N")
'   CAL_PAR  paradas por periodo ("dd/mm/aaaa hh:mm;dd/mm/aaaa hh:mm") = bloqueio de apenas um horario
'   OEE_PADRAO
' Bloqueios sao gravados na hora (Bloquear / Remover). Semana e OEE: botao Salvar.
'==========================================================

Private WithEvents d01 As MSForms.CommandButton
Private WithEvents d02 As MSForms.CommandButton
Private WithEvents d03 As MSForms.CommandButton
Private WithEvents d04 As MSForms.CommandButton
Private WithEvents d05 As MSForms.CommandButton
Private WithEvents d06 As MSForms.CommandButton
Private WithEvents d07 As MSForms.CommandButton
Private WithEvents d08 As MSForms.CommandButton
Private WithEvents d09 As MSForms.CommandButton
Private WithEvents d10 As MSForms.CommandButton
Private WithEvents d11 As MSForms.CommandButton
Private WithEvents d12 As MSForms.CommandButton
Private WithEvents d13 As MSForms.CommandButton
Private WithEvents d14 As MSForms.CommandButton
Private WithEvents d15 As MSForms.CommandButton
Private WithEvents d16 As MSForms.CommandButton
Private WithEvents d17 As MSForms.CommandButton
Private WithEvents d18 As MSForms.CommandButton
Private WithEvents d19 As MSForms.CommandButton
Private WithEvents d20 As MSForms.CommandButton
Private WithEvents d21 As MSForms.CommandButton
Private WithEvents d22 As MSForms.CommandButton
Private WithEvents d23 As MSForms.CommandButton
Private WithEvents d24 As MSForms.CommandButton
Private WithEvents d25 As MSForms.CommandButton
Private WithEvents d26 As MSForms.CommandButton
Private WithEvents d27 As MSForms.CommandButton
Private WithEvents d28 As MSForms.CommandButton
Private WithEvents d29 As MSForms.CommandButton
Private WithEvents d30 As MSForms.CommandButton
Private WithEvents d31 As MSForms.CommandButton
Private WithEvents d32 As MSForms.CommandButton
Private WithEvents d33 As MSForms.CommandButton
Private WithEvents d34 As MSForms.CommandButton
Private WithEvents d35 As MSForms.CommandButton
Private WithEvents d36 As MSForms.CommandButton
Private WithEvents d37 As MSForms.CommandButton
Private WithEvents d38 As MSForms.CommandButton
Private WithEvents d39 As MSForms.CommandButton
Private WithEvents d40 As MSForms.CommandButton
Private WithEvents d41 As MSForms.CommandButton
Private WithEvents d42 As MSForms.CommandButton

Private cel(1 To 42) As MSForms.CommandButton
Private chkDia(1 To 7) As MSForms.CheckBox
Private txtI(1 To 7) As MSForms.TextBox
Private txtF(1 To 7) As MSForms.TextBox
Private txtOEE As MSForms.TextBox
Private txtDe As MSForms.TextBox
Private txtAte As MSForms.TextBox
Private txtH1 As MSForms.TextBox
Private txtH2 As MSForms.TextBox
Private lblMes As MSForms.Label
Private lblMsg As MSForms.Label
Private lstBloq As MSForms.ListBox
Private WithEvents cmdPrev As MSForms.CommandButton
Private WithEvents cmdNext As MSForms.CommandButton
Private WithEvents cmdAdd As MSForms.CommandButton
Private WithEvents cmdRem As MSForms.CommandButton
Private WithEvents cmdSalvar As MSForms.CommandButton
Private WithEvents cmdFechar As MSForms.CommandButton
Private WithEvents optDia As MSForms.OptionButton
Private WithEvents optHor As MSForms.OptionButton

Public Alterou As Boolean         ' True = algum bloqueio/semana/OEE foi gravado (quem abriu o formulario atualiza o planejamento)

Private mAno As Long
Private mMes As Long
Private mOff As Long              ' celulas em branco antes do dia 1
Private mSelIni As Double         ' 0 = sem selecao
Private mSelFim As Double         ' 0 = so um dia
Private mNE As Long
Private mED() As Double
Private mES() As String           ' "S" liberada, "N" bloqueada
Private mEH() As String           ' "hh:nn;hh:nn" (so quando liberada com horario proprio)
Private mNP As Long
Private mPI() As Double
Private mPF() As Double
Private mNL As Long
Private mLK() As String           ' "E" excecao ou "P" parada
Private mLI() As Long

'----------------------------------------------------------
' Textos (sem depender da configuracao regional do Windows)
'----------------------------------------------------------
Private Function Z2(ByVal n As Long) As String
    Z2 = Right$("0" & CStr(n), 2)
End Function

Private Function Dt(ByVal d As Double) As String
    Dim x As Date
    x = CDate(Int(d + 0.0000005))
    Dt = Z2(Day(x)) & "/" & Z2(Month(x)) & "/" & CStr(Year(x))
End Function

' hh:nn de uma fracao de dia (1 = 24:00)
Private Function Hr(ByVal f As Double) As String
    Dim m As Long
    m = CLng(Round(f * 1440#, 0))
    If m >= 1440 Then Hr = "24:00": Exit Function
    Hr = Z2(m \ 60) & ":" & Z2(m Mod 60)
End Function

Private Function DtHr(ByVal v As Double) As String
    DtHr = Dt(v) & " " & Hr(v - Int(v + 0.0000005))
End Function

Private Function Dow3(ByVal d As Double) As String
    Select Case Weekday(CDate(Int(d + 0.0000005)), vbMonday)
        Case 1: Dow3 = "seg"
        Case 2: Dow3 = "ter"
        Case 3: Dow3 = "qua"
        Case 4: Dow3 = "qui"
        Case 5: Dow3 = "sex"
        Case 6: Dow3 = "s" & ChrW(225) & "b"
        Case Else: Dow3 = "dom"
    End Select
End Function

Private Function NomeDia(ByVal wd As Long) As String
    Select Case wd
        Case 1: NomeDia = "Segunda-feira"
        Case 2: NomeDia = "Ter" & ChrW(231) & "a-feira"
        Case 3: NomeDia = "Quarta-feira"
        Case 4: NomeDia = "Quinta-feira"
        Case 5: NomeDia = "Sexta-feira"
        Case 6: NomeDia = "S" & ChrW(225) & "bado"
        Case Else: NomeDia = "Domingo"
    End Select
End Function

Private Function NomeMes(ByVal m As Long) As String
    Select Case m
        Case 1: NomeMes = "Janeiro"
        Case 2: NomeMes = "Fevereiro"
        Case 3: NomeMes = "Mar" & ChrW(231) & "o"
        Case 4: NomeMes = "Abril"
        Case 5: NomeMes = "Maio"
        Case 6: NomeMes = "Junho"
        Case 7: NomeMes = "Julho"
        Case 8: NomeMes = "Agosto"
        Case 9: NomeMes = "Setembro"
        Case 10: NomeMes = "Outubro"
        Case 11: NomeMes = "Novembro"
        Case Else: NomeMes = "Dezembro"
    End Select
End Function

' Fim de horario: "24:00" (ou 00:00) = meia-noite = 1.0. Devolve -1 se invalido.
Private Function ParaFim(ByVal t As String) As Double
    Dim v As Double
    t = Trim$(t)
    If t = "24:00" Then ParaFim = 1#: Exit Function
    v = APS_ParaHora(t)
    If v = 0 Then v = 1#
    ParaFim = v
End Function

Private Sub Msg(ByVal txt As String, ByVal erro As Boolean)
    lblMsg.Caption = txt
    If erro Then lblMsg.ForeColor = RGB(192, 0, 0) Else lblMsg.ForeColor = RGB(0, 110, 60)
End Sub

'----------------------------------------------------------
' Montagem do formulario
'----------------------------------------------------------
Private Function Lb(ByVal nome As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim L As MSForms.Label
    Set L = Me.Controls.Add("Forms.Label.1", nome, True)
    L.Left = x: L.Top = y: L.Width = w: L.Height = h
    L.Caption = txt
    L.Font.Size = 9
    L.WordWrap = True
    Set Lb = L
End Function

Private Function Tx(ByVal nome As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.Add("Forms.TextBox.1", nome, True)
    t.Left = x: t.Top = y: t.Width = w: t.Height = h
    t.Font.Size = 9
    Set Tx = t
End Function

Private Function Bt(ByVal nome As String, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", nome, True)
    b.Left = x: b.Top = y: b.Width = w: b.Height = h
    b.Caption = txt
    b.Font.Size = 9
    Set Bt = b
End Function

Private Sub UserForm_Initialize()
    Dim i As Long, y As Single, L As MSForms.Label, itens() As String, p() As String, d As Long

    Me.Caption = "Calend" & ChrW(225) & "rio de produ" & ChrW(231) & ChrW(227) & "o"
    Me.Width = 560
    Me.Height = 396
    Me.StartUpPosition = 1

    '--- esquerda: semana padrao e OEE -----------------------------------
    Set L = Lb("lbT1", "Semana padr" & ChrW(227) & "o", 10, 6, 230, 16)
    L.Font.Bold = True: L.Font.Size = 10
    Set L = Lb("lbH1", "Dia", 10, 26, 100, 12)
    L.ForeColor = RGB(110, 110, 110)
    Set L = Lb("lbH2", "Das", 120, 26, 44, 12)
    L.ForeColor = RGB(110, 110, 110)
    Set L = Lb("lbH3", "At" & ChrW(233), 170, 26, 44, 12)
    L.ForeColor = RGB(110, 110, 110)
    y = 40
    For i = 1 To 7
        Set chkDia(i) = Me.Controls.Add("Forms.CheckBox.1", "chkDia" & i, True)
        chkDia(i).Left = 10: chkDia(i).Top = y + 1: chkDia(i).Width = 106: chkDia(i).Height = 16
        chkDia(i).Caption = NomeDia(i)
        chkDia(i).Font.Size = 9
        Set txtI(i) = Tx("txtI" & i, 120, y, 44, 17)
        Set txtF(i) = Tx("txtF" & i, 170, y, 44, 17)
        y = y + 21
    Next i
    Set L = Lb("lbOEE", "OEE padr" & ChrW(227) & "o (%)", 10, 192, 100, 14)
    Set txtOEE = Tx("txtOEE", 120, 190, 44, 17)
    Set L = Lb("lbNota", "Hor" & ChrW(225) & "rio de cada dia " & ChrW(233) & " independente (24:00 = meia-noite). Semana e OEE s" & ChrW(243) & " valem depois de Salvar. " & _
               "Os bloqueios ao lado s" & ChrW(227) & "o gravados na hora.", 10, 216, 230, 50)
    L.ForeColor = RGB(110, 110, 110)
    L.Font.Size = 8

    ' valores atuais (CAL_SEM): dia;S/N;inicio;fim
    For i = 1 To 7
        chkDia(i).Value = (i <= 5)
        txtI(i).Text = "06:00"
        txtF(i).Text = "22:00"
    Next i
    itens = Split(APS_CalSemanaTexto(), "|")
    For i = 0 To UBound(itens)
        p = Split(itens(i), ";")
        If UBound(p) >= 1 Then
            d = Val(p(0))
            If d >= 1 And d <= 7 Then
                chkDia(d).Value = (UCase$(Left$(p(1), 1)) = "S")
                If UBound(p) >= 3 Then
                    If APS_ParaHora(p(2)) >= 0 And APS_ParaHora(p(3)) >= 0 Then
                        txtI(d).Text = Hr(APS_ParaHora(p(2)))
                        txtF(d).Text = Hr(ParaFim(p(3)))
                    End If
                End If
            End If
        End If
    Next i
    txtOEE.Text = CStr(Round(APS_OEEPadrao() * 100#, 1))

    '--- direita: bloqueios ----------------------------------------------
    Set L = Lb("lbT2", "Bloquear datas e hor" & ChrW(225) & "rios", 256, 6, 290, 16)
    L.Font.Bold = True: L.Font.Size = 10
    Set cmdPrev = Bt("cmdPrev", "<", 256, 24, 24, 18)
    Set lblMes = Lb("lblMes", "", 282, 27, 228, 14)
    lblMes.Font.Bold = True
    lblMes.TextAlign = fmTextAlignCenter
    Set cmdNext = Bt("cmdNext", ">", 512, 24, 24, 18)
    For i = 1 To 7
        Set L = Lb("lbD" & i, Left$(NomeDia(i), 3), 256 + (i - 1) * 40, 46, 40, 12)
        L.TextAlign = fmTextAlignCenter
        L.Font.Bold = True
        L.Font.Size = 8
    Next i

    Set d01 = NovaCel(1)
    Set d02 = NovaCel(2)
    Set d03 = NovaCel(3)
    Set d04 = NovaCel(4)
    Set d05 = NovaCel(5)
    Set d06 = NovaCel(6)
    Set d07 = NovaCel(7)
    Set d08 = NovaCel(8)
    Set d09 = NovaCel(9)
    Set d10 = NovaCel(10)
    Set d11 = NovaCel(11)
    Set d12 = NovaCel(12)
    Set d13 = NovaCel(13)
    Set d14 = NovaCel(14)
    Set d15 = NovaCel(15)
    Set d16 = NovaCel(16)
    Set d17 = NovaCel(17)
    Set d18 = NovaCel(18)
    Set d19 = NovaCel(19)
    Set d20 = NovaCel(20)
    Set d21 = NovaCel(21)
    Set d22 = NovaCel(22)
    Set d23 = NovaCel(23)
    Set d24 = NovaCel(24)
    Set d25 = NovaCel(25)
    Set d26 = NovaCel(26)
    Set d27 = NovaCel(27)
    Set d28 = NovaCel(28)
    Set d29 = NovaCel(29)
    Set d30 = NovaCel(30)
    Set d31 = NovaCel(31)
    Set d32 = NovaCel(32)
    Set d33 = NovaCel(33)
    Set d34 = NovaCel(34)
    Set d35 = NovaCel(35)
    Set d36 = NovaCel(36)
    Set d37 = NovaCel(37)
    Set d38 = NovaCel(38)
    Set d39 = NovaCel(39)
    Set d40 = NovaCel(40)
    Set d41 = NovaCel(41)
    Set d42 = NovaCel(42)

    Set L = Lb("lbLeg", "vermelho = dia bloqueado   amarelo = hor" & ChrW(225) & "rio bloqueado   azul = sele" & ChrW(231) & ChrW(227) & "o", 256, 172, 290, 12)
    L.Font.Size = 8
    L.ForeColor = RGB(110, 110, 110)

    Lb "lbDe", "De", 256, 192, 18, 14
    Set txtDe = Tx("txtDe", 274, 189, 68, 18)
    Lb "lbAte", "at" & ChrW(233), 350, 192, 18, 14
    Set txtAte = Tx("txtAte", 370, 189, 68, 18)
    Set L = Lb("lbVaz", "(vazio = um dia)", 444, 192, 96, 14)
    L.Font.Size = 8
    L.ForeColor = RGB(110, 110, 110)

    Set optDia = Me.Controls.Add("Forms.OptionButton.1", "optDia", True)
    optDia.Left = 256: optDia.Top = 214: optDia.Width = 78: optDia.Height = 16
    optDia.Caption = "Dia inteiro": optDia.Font.Size = 9
    Set optHor = Me.Controls.Add("Forms.OptionButton.1", "optHor", True)
    optHor.Left = 338: optHor.Top = 214: optHor.Width = 66: optHor.Height = 16
    optHor.Caption = "Hor" & ChrW(225) & "rio": optHor.Font.Size = 9
    Set txtH1 = Tx("txtH1", 408, 212, 40, 18)
    Lb "lbH", ChrW(224) & "s", 452, 215, 14, 14
    Set txtH2 = Tx("txtH2", 468, 212, 40, 18)
    txtH1.Text = "10:00"
    txtH2.Text = "12:00"
    optDia.Value = True

    Set cmdAdd = Bt("cmdAdd", "Bloquear", 256, 238, 110, 24)
    cmdAdd.Font.Bold = True
    Set lblMsg = Lb("lblMsg", "", 372, 236, 170, 30)
    lblMsg.Font.Size = 8

    Set lstBloq = Me.Controls.Add("Forms.ListBox.1", "lstBloq", True)
    lstBloq.Left = 256: lstBloq.Top = 270: lstBloq.Width = 286: lstBloq.Height = 62
    lstBloq.Font.Size = 9
    lstBloq.MultiSelect = fmMultiSelectExtended

    Set cmdRem = Bt("cmdRem", "Remover selecionados", 256, 338, 130, 22)
    Set cmdSalvar = Bt("cmdSalvar", "Salvar semana e OEE", 10, 338, 130, 22)
    cmdSalvar.Font.Bold = True
    Set cmdFechar = Bt("cmdFechar", "Fechar", 146, 338, 70, 22)

    mAno = Year(Date): mMes = Month(Date)
    mSelIni = 0: mSelFim = 0
    CarregarBloqueios
    Pintar
    PreencherLista
End Sub

' o X da janela apenas esconde (quem abriu o formulario le .Alterou e o descarrega)
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

'----------------------------------------------------------
' Leitura / gravacao dos bloqueios (CAL_EXC e CAL_PAR)
'----------------------------------------------------------
Private Sub CarregarBloqueios()
    Dim s As String, itens() As String, p() As String, i As Long, h1 As Double, h2 As Double, d As Double

    mNE = 0
    ReDim mED(1 To 1): ReDim mES(1 To 1): ReDim mEH(1 To 1)
    s = APS_CfgLer("CAL_EXC", "")
    If Len(s) > 0 Then
        itens = Split(s, "|")
        ReDim mED(1 To UBound(itens) + 1): ReDim mES(1 To UBound(itens) + 1): ReDim mEH(1 To UBound(itens) + 1)
        For i = 0 To UBound(itens)
            p = Split(itens(i), ";")
            If UBound(p) >= 1 Then
                d = APS_ParaData(p(0))
                If d >= 0 Then
                    mNE = mNE + 1
                    mED(mNE) = d
                    If UCase$(Left$(Trim$(p(1)), 1)) = "S" Then mES(mNE) = "S" Else mES(mNE) = "N"
                    mEH(mNE) = ""
                    If UBound(p) >= 3 Then mEH(mNE) = Trim$(p(2)) & ";" & Trim$(p(3))
                End If
            End If
        Next i
    End If

    mNP = 0
    ReDim mPI(1 To 1): ReDim mPF(1 To 1)
    s = APS_CfgLer("CAL_PAR", "")
    If Len(s) > 0 Then
        itens = Split(s, "|")
        ReDim mPI(1 To UBound(itens) + 1): ReDim mPF(1 To UBound(itens) + 1)
        For i = 0 To UBound(itens)
            p = Split(itens(i), ";")
            If UBound(p) >= 1 Then
                h1 = APS_ParaDataHora(p(0))
                h2 = APS_ParaDataHora(p(1))
                If h1 >= 0 And h2 > h1 Then
                    mNP = mNP + 1
                    mPI(mNP) = h1
                    mPF(mNP) = h2
                End If
            End If
        Next i
    End If
End Sub

Private Sub GravarBloqueios()
    Dim i As Long, j As Long, t As Long, ord() As Long, sE As String, sP As String

    If mNE > 0 Then
        ReDim ord(1 To mNE)
        For i = 1 To mNE: ord(i) = i: Next i
        For i = 1 To mNE - 1
            For j = i + 1 To mNE
                If mED(ord(j)) < mED(ord(i)) Then t = ord(i): ord(i) = ord(j): ord(j) = t
            Next j
        Next i
        For i = 1 To mNE
            j = ord(i)
            If Len(sE) > 0 Then sE = sE & "|"
            sE = sE & Dt(mED(j)) & ";" & mES(j)
            If mES(j) = "S" And Len(mEH(j)) > 0 Then sE = sE & ";" & mEH(j)
        Next i
    End If
    If mNP > 0 Then
        ReDim ord(1 To mNP)
        For i = 1 To mNP: ord(i) = i: Next i
        For i = 1 To mNP - 1
            For j = i + 1 To mNP
                If mPI(ord(j)) < mPI(ord(i)) Then t = ord(i): ord(i) = ord(j): ord(j) = t
            Next j
        Next i
        For i = 1 To mNP
            j = ord(i)
            If Len(sP) > 0 Then sP = sP & "|"
            sP = sP & DtHr(mPI(j)) & ";" & DtHr(mPF(j))
        Next i
    End If
    APS_CfgGravar "CAL_EXC", sE
    APS_CfgGravar "CAL_PAR", sP
    APS_CalRecarregar
    Alterou = True
End Sub

Private Function IdxExc(ByVal d As Double) As Long
    Dim i As Long
    For i = 1 To mNE
        If mED(i) = d Then IdxExc = i: Exit Function
    Next i
End Function

Private Sub PoeDiaInteiro(ByVal d As Double)
    Dim k As Long
    k = IdxExc(d)
    If k = 0 Then
        mNE = mNE + 1
        If mNE > UBound(mED) Then ReDim Preserve mED(1 To mNE + 20): ReDim Preserve mES(1 To mNE + 20): ReDim Preserve mEH(1 To mNE + 20)
        k = mNE
    End If
    mED(k) = d: mES(k) = "N": mEH(k) = ""
End Sub

Private Sub PoeHorario(ByVal d As Double, ByVal h1 As Double, ByVal h2 As Double)
    Dim i As Long, a As Double, b As Double
    a = APS_ArredMin(d + h1)
    b = APS_ArredMin(d + h2)
    For i = 1 To mNP
        If Abs(mPI(i) - a) < 0.0000005 And Abs(mPF(i) - b) < 0.0000005 Then Exit Sub
    Next i
    mNP = mNP + 1
    If mNP > UBound(mPI) Then ReDim Preserve mPI(1 To mNP + 20): ReDim Preserve mPF(1 To mNP + 20)
    mPI(mNP) = a
    mPF(mNP) = b
End Sub

'----------------------------------------------------------
' Calendario do mes (cores = situacao de cada dia)
'----------------------------------------------------------
Private Sub Pintar()
    Dim i As Long, dia As Long, nd As Long, d As Double, k As Long, p As Long
    Dim cheio As Boolean, parc As Boolean, sel As Boolean, selA As Double, selB As Double, wd As Long

    lblMes.Caption = NomeMes(mMes) & " de " & mAno
    mOff = Weekday(DateSerial(mAno, mMes, 1), vbMonday) - 1
    nd = Day(DateSerial(mAno, mMes + 1, 0))
    selA = mSelIni
    If mSelFim > 0 Then selB = mSelFim Else selB = mSelIni
    For i = 1 To 42
        dia = i - mOff
        If dia < 1 Or dia > nd Then
            cel(i).Caption = ""
            cel(i).Enabled = False
            cel(i).BackColor = RGB(255, 255, 255)
        Else
            d = CDbl(DateSerial(mAno, mMes, dia))
            cel(i).Caption = CStr(dia)
            cel(i).Enabled = True
            cheio = False: parc = False
            k = IdxExc(d)
            If k > 0 Then
                If mES(k) = "N" Then cheio = True Else parc = True
            End If
            For p = 1 To mNP
                If mPI(p) < d + 1# And mPF(p) > d Then parc = True
            Next p
            sel = (selA > 0 And d >= selA And d <= selB)
            wd = Weekday(CDate(d), vbMonday)
            cel(i).Font.Bold = (d = CDbl(Date))
            If sel Then
                cel(i).BackColor = RGB(31, 78, 120)
                cel(i).ForeColor = RGB(255, 255, 255)
            ElseIf cheio Then
                cel(i).BackColor = RGB(244, 199, 195)
                cel(i).ForeColor = RGB(150, 0, 0)
            ElseIf parc Then
                cel(i).BackColor = RGB(255, 235, 156)
                cel(i).ForeColor = RGB(0, 0, 0)
            Else
                cel(i).BackColor = RGB(255, 255, 255)
                If chkDia(wd).Value Then cel(i).ForeColor = RGB(0, 0, 0) Else cel(i).ForeColor = RGB(150, 150, 150)
            End If
        End If
    Next i
End Sub

Private Sub DiaClick(ByVal i As Long)
    Dim dia As Long, d As Double
    dia = i - mOff
    If dia < 1 Then Exit Sub
    d = CDbl(DateSerial(mAno, mMes, dia))
    If mSelIni = 0 Or mSelFim > 0 Then
        mSelIni = d: mSelFim = 0
        txtDe.Text = Dt(d): txtAte.Text = ""
    Else
        If d < mSelIni Then
            mSelFim = mSelIni: mSelIni = d
        ElseIf d > mSelIni Then
            mSelFim = d
        End If
        txtDe.Text = Dt(mSelIni)
        If mSelFim > 0 Then txtAte.Text = Dt(mSelFim) Else txtAte.Text = ""
    End If
    Pintar
End Sub

Private Sub cmdPrev_Click()
    mMes = mMes - 1
    If mMes < 1 Then mMes = 12: mAno = mAno - 1
    Pintar
End Sub

Private Sub cmdNext_Click()
    mMes = mMes + 1
    If mMes > 12 Then mMes = 1: mAno = mAno + 1
    Pintar
End Sub

Private Sub optDia_Click()
    txtH1.Enabled = False: txtH2.Enabled = False
End Sub

Private Sub optHor_Click()
    txtH1.Enabled = True: txtH2.Enabled = True
End Sub

'----------------------------------------------------------
' Lista de bloqueios cadastrados
'----------------------------------------------------------
Private Sub PreencherLista()
    Dim i As Long, j As Long, n As Long, key() As Double, k As String, ix As Long, t As Long, txt As String, tmpK As String, tmpI As Long, tmpD As Double

    n = mNE + mNP
    lstBloq.Clear
    mNL = 0
    If n = 0 Then Exit Sub
    ReDim mLK(1 To n): ReDim mLI(1 To n): ReDim key(1 To n)
    For i = 1 To mNE
        mNL = mNL + 1: mLK(mNL) = "E": mLI(mNL) = i: key(mNL) = mED(i)
    Next i
    For i = 1 To mNP
        mNL = mNL + 1: mLK(mNL) = "P": mLI(mNL) = i: key(mNL) = mPI(i)
    Next i
    For i = 1 To mNL - 1
        For j = i + 1 To mNL
            If key(j) < key(i) Then
                tmpD = key(i): key(i) = key(j): key(j) = tmpD
                tmpK = mLK(i): mLK(i) = mLK(j): mLK(j) = tmpK
                tmpI = mLI(i): mLI(i) = mLI(j): mLI(j) = tmpI
            End If
        Next j
    Next i
    For i = 1 To mNL
        ix = mLI(i)
        If mLK(i) = "E" Then
            txt = Dt(mED(ix)) & " (" & Dow3(mED(ix)) & ")  -  "
            If mES(ix) = "N" Then
                txt = txt & "dia inteiro"
            ElseIf Len(mEH(ix)) > 0 Then
                txt = txt & "trabalha " & Replace(mEH(ix), ";", " " & ChrW(224) & "s ")
            Else
                txt = txt & "dia liberado"
            End If
        Else
            If Int(mPI(ix) + 0.0000005) = Int(mPF(ix) - 0.0000005) Then
                txt = Dt(mPI(ix)) & " (" & Dow3(mPI(ix)) & ")  -  " & Hr(mPI(ix) - Int(mPI(ix) + 0.0000005)) & " " & ChrW(224) & "s " & Hr(mPF(ix) - Int(mPI(ix) + 0.0000005))
            Else
                txt = DtHr(mPI(ix)) & "  at" & ChrW(233) & "  " & DtHr(mPF(ix))
            End If
        End If
        lstBloq.AddItem txt
    Next i
End Sub

'----------------------------------------------------------
' Ordens ja programadas dentro do que foi bloqueado (so aviso; nada e movido aqui)
'----------------------------------------------------------
Private Function OrdensNoBloqueio(ByVal d1 As Double, ByVal d2 As Double, ByVal horario As Boolean, ByVal h1 As Double, ByVal h2 As Double) As Long
    Dim ops() As tOperacao, n As Long, i As Long, d As Double, a As Double, b As Double, achou As Boolean, c As Long

    On Error GoTo Sai
    n = APS_LerOps(ops)
    If n <= 0 Then Exit Function
    For i = 1 To n
        If Len(ops(i).tipoAux) = 0 Then
            If APS_Ativa(ops(i)) Then
                achou = False
                If horario Then
                    For d = d1 To d2
                        a = d + h1: b = d + h2
                        If ops(i).ini < b - 0.0000005 And ops(i).fim > a + 0.0000005 Then achou = True: Exit For
                    Next d
                Else
                    If ops(i).ini < d2 + 1# - 0.0000005 And ops(i).fim > d1 + 0.0000005 Then achou = True
                End If
                If achou Then c = c + 1
            End If
        End If
    Next i
    OrdensNoBloqueio = c
Sai:
End Function

'----------------------------------------------------------
' Bloquear / remover
'----------------------------------------------------------
Private Sub cmdAdd_Click()
    Dim d1 As Double, d2 As Double, d As Double, h1 As Double, h2 As Double, t As Double, nd As Long, nOrd As Long, txt As String

    d1 = APS_ParaData(txtDe.Text)
    If d1 < 0 Then Msg "Informe a data (dd/mm/aaaa) ou clique no calend" & ChrW(225) & "rio.", True: Exit Sub
    If Len(Trim$(txtAte.Text)) = 0 Then
        d2 = d1
    Else
        d2 = APS_ParaData(txtAte.Text)
        If d2 < 0 Then Msg "Data final inv" & ChrW(225) & "lida (dd/mm/aaaa).", True: Exit Sub
    End If
    If d2 < d1 Then t = d1: d1 = d2: d2 = t
    If d2 - d1 > 366 Then Msg "Per" & ChrW(237) & "odo longo demais (m" & ChrW(225) & "ximo 366 dias).", True: Exit Sub

    If optHor.Value Then
        h1 = APS_ParaHora(txtH1.Text)
        h2 = ParaFim(txtH2.Text)
        If h1 < 0 Or h2 < 0 Then Msg "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido (use hh:mm).", True: Exit Sub
        If h2 <= h1 Then Msg "O hor" & ChrW(225) & "rio final deve ser maior que o inicial.", True: Exit Sub
        For d = d1 To d2
            PoeHorario d, h1, h2
            nd = nd + 1
        Next d
    Else
        For d = d1 To d2
            PoeDiaInteiro d
            nd = nd + 1
        Next d
    End If
    GravarBloqueios
    nOrd = OrdensNoBloqueio(d1, d2, optHor.Value, h1, h2)

    If optHor.Value Then
        txt = "Bloqueado " & Hr(h1) & " " & ChrW(224) & "s " & Hr(h2) & " em " & nd & " dia(s)."
    Else
        txt = "Dia inteiro bloqueado em " & nd & " dia(s)."
    End If
    If nOrd > 0 Then
        txt = txt & " Aten" & ChrW(231) & ChrW(227) & "o: " & nOrd & " ordem(ns) j" & ChrW(225) & " programada(s) nesse per" & ChrW(237) & "odo " & ChrW(8212) & " confira no planejamento."
        Msg txt, True
    Else
        Msg txt, False
    End If
    mSelIni = 0: mSelFim = 0
    Pintar
    PreencherLista
End Sub

Private Sub cmdRem_Click()
    Dim i As Long, j As Long, nSel As Long, delE() As Boolean, delP() As Boolean, k As Long
    Dim nED As Long, nPD As Long

    If lstBloq.ListCount = 0 Then Exit Sub
    ReDim delE(1 To IIf(mNE > 0, mNE, 1)): ReDim delP(1 To IIf(mNP > 0, mNP, 1))
    For i = 0 To lstBloq.ListCount - 1
        If lstBloq.Selected(i) Then
            nSel = nSel + 1
            If mLK(i + 1) = "E" Then delE(mLI(i + 1)) = True Else delP(mLI(i + 1)) = True
        End If
    Next i
    If nSel = 0 Then Msg "Selecione na lista o que deseja remover.", True: Exit Sub

    k = 0
    For i = 1 To mNE
        If Not delE(i) Then
            k = k + 1
            mED(k) = mED(i): mES(k) = mES(i): mEH(k) = mEH(i)
        End If
    Next i
    mNE = k
    k = 0
    For i = 1 To mNP
        If Not delP(i) Then
            k = k + 1
            mPI(k) = mPI(i): mPF(k) = mPF(i)
        End If
    Next i
    mNP = k
    GravarBloqueios
    Msg nSel & " item(ns) removido(s).", False
    Pintar
    PreencherLista
End Sub

'----------------------------------------------------------
' Semana padrao e OEE
'----------------------------------------------------------
Private Sub cmdSalvar_Click()
    Dim i As Long, s As String, h1 As Double, h2 As Double, oee As Double

    For i = 1 To 7
        h1 = APS_ParaHora(txtI(i).Text)
        h2 = ParaFim(txtF(i).Text)
        If chkDia(i).Value Then
            If h1 < 0 Or h2 < 0 Then
                Msg "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido em " & NomeDia(i) & " (use hh:mm).", True
                Exit Sub
            End If
            If h2 <= h1 Then
                Msg "O fim deve ser maior que o in" & ChrW(237) & "cio em " & NomeDia(i) & ".", True
                Exit Sub
            End If
        Else
            If h1 < 0 Or h2 < 0 Or h2 <= h1 Then h1 = 6# / 24#: h2 = 22# / 24#
        End If
        s = s & i & ";"
        If chkDia(i).Value Then s = s & "S;" Else s = s & "N;"
        s = s & Hr(h1) & ";"
        If h2 >= 1# Then s = s & "00:00|" Else s = s & Hr(h2) & "|"
    Next i
    s = Left$(s, Len(s) - 1)

    oee = Val(Replace(Replace(Trim$(txtOEE.Text), "%", ""), ",", "."))
    If oee <= 0 Or oee > 100 Then
        Msg "OEE padr" & ChrW(227) & "o deve estar entre 1 e 100.", True
        Exit Sub
    End If

    APS_CfgGravar "CAL_SEM", s
    APS_CfgGravar "OEE_PADRAO", Replace(CStr(oee), ",", ".")
    APS_CalRecarregar
    Alterou = True
    Me.Hide
End Sub

Private Sub cmdFechar_Click()
    Me.Hide
End Sub

Private Function NovaCel(ByVal i As Long) As MSForms.CommandButton
    Dim b As MSForms.CommandButton, r As Long, c As Long
    r = (i - 1) \ 7
    c = (i - 1) Mod 7
    Set b = Bt("dia" & i, "", 256 + c * 40, 62 + r * 18, 40, 17)
    b.Font.Size = 8
    Set cel(i) = b
    Set NovaCel = b
End Function
Private Sub d01_Click()
    DiaClick 1
End Sub

Private Sub d02_Click()
    DiaClick 2
End Sub

Private Sub d03_Click()
    DiaClick 3
End Sub

Private Sub d04_Click()
    DiaClick 4
End Sub

Private Sub d05_Click()
    DiaClick 5
End Sub

Private Sub d06_Click()
    DiaClick 6
End Sub

Private Sub d07_Click()
    DiaClick 7
End Sub

Private Sub d08_Click()
    DiaClick 8
End Sub

Private Sub d09_Click()
    DiaClick 9
End Sub

Private Sub d10_Click()
    DiaClick 10
End Sub

Private Sub d11_Click()
    DiaClick 11
End Sub

Private Sub d12_Click()
    DiaClick 12
End Sub

Private Sub d13_Click()
    DiaClick 13
End Sub

Private Sub d14_Click()
    DiaClick 14
End Sub

Private Sub d15_Click()
    DiaClick 15
End Sub

Private Sub d16_Click()
    DiaClick 16
End Sub

Private Sub d17_Click()
    DiaClick 17
End Sub

Private Sub d18_Click()
    DiaClick 18
End Sub

Private Sub d19_Click()
    DiaClick 19
End Sub

Private Sub d20_Click()
    DiaClick 20
End Sub

Private Sub d21_Click()
    DiaClick 21
End Sub

Private Sub d22_Click()
    DiaClick 22
End Sub

Private Sub d23_Click()
    DiaClick 23
End Sub

Private Sub d24_Click()
    DiaClick 24
End Sub

Private Sub d25_Click()
    DiaClick 25
End Sub

Private Sub d26_Click()
    DiaClick 26
End Sub

Private Sub d27_Click()
    DiaClick 27
End Sub

Private Sub d28_Click()
    DiaClick 28
End Sub

Private Sub d29_Click()
    DiaClick 29
End Sub

Private Sub d30_Click()
    DiaClick 30
End Sub

Private Sub d31_Click()
    DiaClick 31
End Sub

Private Sub d32_Click()
    DiaClick 32
End Sub

Private Sub d33_Click()
    DiaClick 33
End Sub

Private Sub d34_Click()
    DiaClick 34
End Sub

Private Sub d35_Click()
    DiaClick 35
End Sub

Private Sub d36_Click()
    DiaClick 36
End Sub

Private Sub d37_Click()
    DiaClick 37
End Sub

Private Sub d38_Click()
    DiaClick 38
End Sub

Private Sub d39_Click()
    DiaClick 39
End Sub

Private Sub d40_Click()
    DiaClick 40
End Sub

Private Sub d41_Click()
    DiaClick 41
End Sub

Private Sub d42_Click()
    DiaClick 42
End Sub

