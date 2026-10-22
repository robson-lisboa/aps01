VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPS_Parametros 
   Caption         =   "Cadastro de Parametros"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8300
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAPS_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' APS PURAN - CADASTRO DE PARAMETROS (tela)
' Fluxo progressivo:  Area -> Maquina (ou etapa) -> Produto -> [Carregar parametros] -> editar -> salvar.
' A tela inicial mostra so a selecao. Depois de carregar, aparecem SOMENTE os campos da selecao atual;
' ao mudar area, maquina ou produto o painel e fechado (com confirmacao se houver alteracao nao salva).
'   Area 1  Producao por maquina : lote padrao, comprimidos por caixa, velocidade (FETTE 2090, FETTE 2, Mediseal, Blisterflex)
'   Area 2  Dados do produto     : quantidade de comprimidos, peso, granulacao, sala de mistura
'   Area 3  Mistura e granulacao : capacidade geral da etapa (kg) e, na Sala de Mistura, meia dosagem do produto
' Campo vazio = parametro nao cadastrado. Salvar grava, RELE o cadastro e confere. Nao cria operacao nem card
' e nao altera operacoes ja gravadas. Dados e regras: modAPS_Parametros (sem alteracao no armazenamento).
'==========================================================

Private WithEvents cboArea As MSForms.ComboBox
Private WithEvents cboMaq As MSForms.ComboBox
Private WithEvents txtBusca As MSForms.TextBox
Private WithEvents lstRes As MSForms.ListBox
Private WithEvents cmdCarregar As MSForms.CommandButton
Private WithEvents cmdSalvar As MSForms.CommandButton
Private WithEvents cmdDesf As MSForms.CommandButton
Private WithEvents cmdFechar As MSForms.CommandButton
Private WithEvents tR1 As MSForms.TextBox
Private WithEvents tR2 As MSForms.TextBox
Private WithEvents tR3 As MSForms.TextBox
Private WithEvents cU1 As MSForms.ComboBox

Private lblMaqL As MSForms.Label
Private lblProdL As MSForms.Label
Private lblSel As MSForms.Label
Private lblMsg As MSForms.Label
Private lblGrp1 As MSForms.Label
Private lblGrp2 As MSForms.Label
Private lblHint As MSForms.Label
Private lblTempoL As MSForms.Label
Private lblTempo As MSForms.Label
Private lblR(1 To 4) As MSForms.Label
Private txtR(1 To 4) As MSForms.TextBox

Private mBloq As Long
Private mModo As Long
Private mLMaq As String
Private mLCod As String
Private mProdPend As String
Private mSnap(1 To 5) As String
Private mRCod() As String
Private mNR As Long

'----------------------------------------------------------
' Montagem
'----------------------------------------------------------
Private Function Lb(ByVal txt As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                    Optional ByVal h As Single = 14, Optional ByVal negrito As Boolean = False) As MSForms.Label
    Dim l As MSForms.Label

    Set l = Me.Controls.Add("Forms.Label.1", , True)
    l.Caption = txt
    l.Left = x
    l.Top = y
    l.Width = w
    l.Height = h
    l.Font.Size = 8
    l.Font.Bold = negrito
    Set Lb = l
End Function

Private Function Tx(ByVal nome As String, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox

    Set t = Me.Controls.Add("Forms.TextBox.1", nome, True)
    t.Left = x
    t.Top = y
    t.Width = w
    t.Height = 18
    t.Font.Size = 9
    Set Tx = t
End Function

Private Function Cb(ByVal nome As String, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.ComboBox
    Dim c As MSForms.ComboBox

    Set c = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    c.Left = x
    c.Top = y
    c.Width = w
    c.Height = 18
    c.Style = 2
    c.Font.Size = 9
    Set Cb = c
End Function

Private Function Bt(ByVal nome As String, ByVal cap As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton

    Set b = Me.Controls.Add("Forms.CommandButton.1", nome, True)
    b.Caption = cap
    b.Left = x
    b.Top = y
    b.Width = w
    b.Height = 24
    b.Font.Size = 8
    Set Bt = b
End Function

Private Sub UserForm_Initialize()
    Dim k As Long, l As MSForms.Label

    Me.Caption = "Cadastro de Par" & ChrW(226) & "metros"
    Me.Width = 656
    Me.Height = 168
    Me.StartUpPosition = 1

    mBloq = mBloq + 1

    Set l = Lb("Cadastro de Par" & ChrW(226) & "metros", 12, 6, 400, 20, True)
    l.Font.Size = 13
    l.ForeColor = RGB(31, 73, 125)
    Set l = Lb("", 12, 30, 616, 1)
    l.BackColor = RGB(190, 200, 215)

    ' ---------- selecao: Area > Maquina/etapa > Produto
    Lb "1. " & ChrW(193) & "rea", 12, 38, 150, 12, True
    Set cboArea = Cb("cboArea", 12, 52, 150)
    cboArea.AddItem "Produ" & ChrW(231) & ChrW(227) & "o por m" & ChrW(225) & "quina"
    cboArea.AddItem "Dados do produto"
    cboArea.AddItem "Mistura e granula" & ChrW(231) & ChrW(227) & "o"

    Set lblMaqL = Lb("2. M" & ChrW(225) & "quina", 172, 38, 150, 12, True)
    Set cboMaq = Cb("cboMaq", 172, 52, 150)
    cboMaq.Enabled = False

    Set lblProdL = Lb("3. Produto (c" & ChrW(243) & "digo ou nome)", 332, 38, 200, 12, True)
    Set txtBusca = Tx("txtBusca", 332, 52, 200)
    txtBusca.Enabled = False

    Set cmdCarregar = Bt("cmdCarregar", "Carregar par" & ChrW(226) & "metros", 540, 49, 88)
    cmdCarregar.Font.Bold = True
    cmdCarregar.Height = 24

    Set lblSel = Lb("", 12, 80, 616, 18)
    lblSel.Font.Size = 9
    lblSel.BackColor = RGB(236, 241, 247)
    lblSel.ForeColor = RGB(100, 100, 100)

    ' ---------- area de edicao (so aparece depois de carregar; os rotulos mudam conforme a selecao)
    Set lblGrp1 = Lb("", 12, 112, 400, 14, True)
    lblGrp1.Font.Size = 9
    lblGrp1.ForeColor = RGB(31, 73, 125)
    Set lblGrp2 = Lb("", 12, 194, 400, 14, True)
    lblGrp2.Font.Size = 9
    lblGrp2.ForeColor = RGB(31, 73, 125)

    For k = 1 To 4
        Set lblR(k) = Lb("", 12, 136 + (k - 1) * 26 + 3, 180)
        Set txtR(k) = Tx("tR" & k, 196, 136 + (k - 1) * 26, 120)
    Next k
    Set tR1 = txtR(1)
    Set tR2 = txtR(2)
    Set tR3 = txtR(3)

    Set cU1 = Cb("cU1", 322, 136, 130)

    Set lblTempoL = Lb("Tempo de refer" & ChrW(234) & "ncia", 12, 217, 180)
    Set lblTempo = Lb("", 196, 217, 200, 14, True)
    lblTempo.ForeColor = RGB(31, 73, 125)
    Set lblHint = Lb("", 12, 238, 616, 14)
    lblHint.ForeColor = RGB(100, 100, 100)

    Set cmdSalvar = Bt("cmdSalvar", "Salvar altera" & ChrW(231) & ChrW(245) & "es", 12, 292, 130)
    cmdSalvar.Font.Bold = True
    Set cmdDesf = Bt("cmdDesf", "Desfazer altera" & ChrW(231) & ChrW(245) & "es", 150, 292, 130)
    Set cmdFechar = Bt("cmdFechar", "Fechar", 538, 292, 90)

    Set lblMsg = Lb("", 12, 324, 616, 30)
    lblMsg.WordWrap = True

    ' lista de resultados da pesquisa: janela flutuante criada por ultimo (fica na frente)
    Set lstRes = Me.Controls.Add("Forms.ListBox.1", "lstRes", True)
    lstRes.Left = 332
    lstRes.Top = 71
    lstRes.Width = 296
    lstRes.Height = 100
    lstRes.ColumnCount = 2
    lstRes.ColumnWidths = "58 pt;220 pt"
    lstRes.Font.Size = 8
    lstRes.Visible = False

    mBloq = mBloq - 1

    Descarregar
End Sub

'----------------------------------------------------------
' Utilitarios
'----------------------------------------------------------
Private Sub Msg(ByVal t As String, ByVal erro As Boolean)
    If erro Then
        lblMsg.ForeColor = RGB(192, 0, 0)
    Else
        lblMsg.ForeColor = RGB(0, 110, 40)
    End If
    lblMsg.Caption = t
End Sub

Private Function ModoSel() As Long
    ModoSel = cboArea.ListIndex + 1
End Function

Private Function NomeProd(ByVal cod As String) As String
    If Len(cod) = 0 Then Exit Function
    NomeProd = cod & " " & ChrW(8212) & " " & APS_NomeDoCodigo(cod)
End Function

Private Sub SelecionarItem(ByVal cbo As MSForms.ComboBox, ByVal txt As String)
    Dim i As Long

    cbo.ListIndex = -1
    For i = 0 To cbo.ListCount - 1
        If APS_Ig(cbo.List(i), txt) Then cbo.ListIndex = i: Exit For
    Next i
End Sub

' maquinas / etapas da area escolhida (lista vem do proprio cadastro de regras: modAPS_Parametros)
Private Sub PreencherMaq(ByVal modo As Long)
    Dim ms() As String, k As Long, i As Long

    mBloq = mBloq + 1
    cboMaq.Clear

    If modo = 1 Then
        lblMaqL.Caption = "2. M" & ChrW(225) & "quina"
        k = APS_RegMaquinas(PAR_LOTE, ms)
    ElseIf modo = 3 Then
        lblMaqL.Caption = "2. Etapa"
        k = APS_RegMaquinas(PAR_CAPAC, ms)
    Else
        lblMaqL.Caption = "2. M" & ChrW(225) & "quina"
        k = 0
    End If

    For i = 1 To k
        cboMaq.AddItem ms(i)
    Next i

    cboMaq.Enabled = (k > 0)
    mBloq = mBloq - 1
End Sub

' o produto e obrigatorio nas areas 1 e 2; na area 3 so serve para a meia dosagem (Sala de Mistura)
Private Sub AjustarProduto()
    Dim modo As Long, ativo As Boolean

    modo = ModoSel()

    ativo = (modo = 1 Or modo = 2)
    If modo = 3 Then ativo = APS_EhSala(cboMaq.Text)

    If modo = 3 Then
        lblProdL.Caption = "3. Produto (opcional: meia dosagem)"
    Else
        lblProdL.Caption = "3. Produto (c" & ChrW(243) & "digo ou nome)"
    End If

    txtBusca.Enabled = ativo

    If Not ativo Then
        mBloq = mBloq + 1
        mProdPend = ""
        txtBusca.Text = ""
        mBloq = mBloq - 1
    End If

    lstRes.Visible = False
End Sub

'----------------------------------------------------------
' Pesquisa de produtos (janela flutuante; so lista produtos, nunca parametros)
'----------------------------------------------------------
Private Sub Pesquisar(ByVal q As String)
    Dim cs() As String, ns() As String, i As Long, k As Long

    k = APS_ParBuscaProduto(q, cs, ns)

    mBloq = mBloq + 1
    lstRes.Clear
    mNR = k
    ReDim mRCod(1 To k + 1)

    For i = 1 To k
        lstRes.AddItem cs(i)
        lstRes.List(i - 1, 1) = ns(i)
        mRCod(i) = cs(i)
    Next i
    mBloq = mBloq - 1

    lstRes.Visible = (k > 0)
    If k = 0 Then Msg "Nenhum produto encontrado.", True
End Sub

Private Sub txtBusca_Enter()
    If mBloq > 0 Then Exit Sub
    If Not txtBusca.Enabled Then Exit Sub
    If Len(mProdPend) = 0 Then Pesquisar txtBusca.Text
End Sub

Private Sub txtBusca_Change()
    If mBloq > 0 Then Exit Sub

    ' digitar nao troca a selecao; so uma escolha na lista troca o produto
    If Len(mProdPend) > 0 Then
        If txtBusca.Text <> NomeProd(mProdPend) Then mProdPend = ""
    End If

    Pesquisar txtBusca.Text
End Sub

Private Sub txtBusca_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        lstRes.Visible = False
    ElseIf KeyCode = 13 Then
        KeyCode = 0
        Pesquisar txtBusca.Text
        If mNR = 1 Then EscolherProduto mRCod(1)
    End If
End Sub

Private Sub lstRes_Click()
    Dim i As Long

    If mBloq > 0 Then Exit Sub

    i = lstRes.ListIndex + 1
    If i < 1 Or i > mNR Then Exit Sub

    EscolherProduto mRCod(i)
End Sub

Private Sub EscolherProduto(ByVal cod As String)
    If Not ConfirmarTroca() Then
        lstRes.Visible = False
        Exit Sub
    End If

    Descarregar

    mBloq = mBloq + 1
    mProdPend = cod
    txtBusca.Text = NomeProd(cod)
    mBloq = mBloq - 1

    lstRes.Visible = False
    Msg "", False
End Sub

'----------------------------------------------------------
' Troca de selecao
'----------------------------------------------------------
' True = pode trocar. Com alteracao nao salva pergunta; se o usuario recusar, a selecao volta ao que esta carregado.
Private Function ConfirmarTroca() As Boolean
    If mModo = 0 Then ConfirmarTroca = True: Exit Function
    If Not TemMudancas() Then ConfirmarTroca = True: Exit Function

    If MsgBox("H" & ChrW(225) & " altera" & ChrW(231) & ChrW(245) & "es n" & ChrW(227) & "o salvas. Descart" & ChrW(225) & "-las e mudar a sele" & ChrW(231) & ChrW(227) & "o?", vbYesNo + vbQuestion, APS_TITULO) = vbYes Then
        ConfirmarTroca = True
    Else
        RestaurarSelecao
    End If
End Function

Private Sub RestaurarSelecao()
    mBloq = mBloq + 1
    cboArea.ListIndex = mModo - 1
    mBloq = mBloq - 1

    PreencherMaq mModo

    mBloq = mBloq + 1
    If Len(mLMaq) > 0 Then SelecionarItem cboMaq, mLMaq
    mProdPend = mLCod
    txtBusca.Text = NomeProd(mLCod)
    mBloq = mBloq - 1

    AjustarProduto
    mBloq = mBloq + 1
    mProdPend = mLCod
    txtBusca.Text = NomeProd(mLCod)
    mBloq = mBloq - 1
End Sub

Private Sub cboArea_Change()
    If mBloq > 0 Then Exit Sub
    If Not ConfirmarTroca() Then Exit Sub

    Descarregar
    PreencherMaq ModoSel()
    AjustarProduto
    Msg "", False
End Sub

Private Sub cboMaq_Change()
    If mBloq > 0 Then Exit Sub
    If Not ConfirmarTroca() Then Exit Sub

    Descarregar
    AjustarProduto
    Msg "", False
End Sub

'----------------------------------------------------------
' Painel de edicao
'----------------------------------------------------------
Private Sub MostrarPainel(ByVal ligado As Boolean)
    Dim k As Long

    lblGrp1.Visible = ligado
    lblGrp2.Visible = False
    lblHint.Visible = ligado
    lblTempoL.Visible = False
    lblTempo.Visible = False
    cU1.Visible = False

    For k = 1 To 4
        lblR(k).Visible = False
        txtR(k).Visible = False
    Next k

    cmdSalvar.Visible = ligado
    cmdDesf.Visible = ligado

    If ligado Then
        cmdFechar.Top = 292
        lblMsg.Top = 324
        lblMsg.Width = 616
        Me.Height = 394
    Else
        cmdFechar.Top = 104
        lblMsg.Top = 102
        lblMsg.Width = 520
        Me.Height = 168
    End If
End Sub

' Fecha o painel e limpa os campos: nada da selecao anterior permanece na tela
Private Sub Descarregar()
    Dim k As Long

    mBloq = mBloq + 1
    mModo = 0
    mLMaq = ""
    mLCod = ""
    For k = 1 To 4
        txtR(k).Text = ""
        lblR(k).Caption = ""
    Next k
    cU1.Clear
    lblTempo.Caption = ""
    lblHint.Caption = ""
    lblGrp1.Caption = ""
    lblGrp2.Caption = ""
    For k = 1 To 5
        mSnap(k) = ""
    Next k
    mBloq = mBloq - 1

    MostrarPainel False
    lblSel.ForeColor = RGB(100, 100, 100)
    lblSel.Caption = "  Escolha a " & ChrW(225) & "rea, a m" & ChrW(225) & "quina e o produto e clique em Carregar par" & ChrW(226) & "metros."
End Sub

Private Sub Linha(ByVal k As Long, ByVal rotulo As String, ByVal topo As Single, ByVal larg As Single)
    lblR(k).Caption = rotulo
    lblR(k).Top = topo + 3
    lblR(k).Visible = True
    txtR(k).Top = topo
    txtR(k).Width = larg
    txtR(k).Visible = True
End Sub

Private Sub UnidadeNaLinha(ByVal topo As Single)
    cU1.Top = topo
    cU1.Left = 322
    cU1.Visible = True
End Sub

Private Sub TirarFoto()
    Dim k As Long

    For k = 1 To 4
        mSnap(k) = Trim$(txtR(k).Text)
    Next k
    mSnap(5) = Trim$(cU1.Text)
End Sub

Private Function TemMudancas() As Boolean
    Dim k As Long

    If mModo = 0 Then Exit Function

    For k = 1 To 4
        If Trim$(txtR(k).Text) <> mSnap(k) Then TemMudancas = True: Exit Function
    Next k
    If Trim$(cU1.Text) <> mSnap(5) Then TemMudancas = True
End Function

'----------------------------------------------------------
' Tempo de referencia (area 1; so informativo; sem OEE)
'----------------------------------------------------------
Private Sub CalcTempo()
    Dim lp As Double, cpc As Double, vel As Double, un As String
    Dim cx As Double, sn As Double, so As Double, porCx As Boolean

    If lblTempo Is Nothing Then Exit Sub
    lblTempo.Caption = ""
    If mModo <> 1 Then Exit Sub
    If Not APS_ParUsaVelocidade(mLMaq) Then Exit Sub

    vel = APS_ParNumero(txtR(3).Text)
    If vel <= 0 Then Exit Sub

    un = Trim$(cU1.Text)
    If Len(un) = 0 Then lblTempo.Caption = "escolha a unidade": Exit Sub

    lp = APS_ParNumero(txtR(1).Text)
    If lp <= 0 Then lblTempo.Caption = "informe o lote": Exit Sub

    porCx = APS_Ig(un, UN_CX_MIN)
    cpc = APS_ParNumero(txtR(2).Text)

    If porCx And cpc <= 0 Then lblTempo.Caption = "falta comprimidos por caixa": Exit Sub

    If APS_CalcTempo(lp, cpc, vel, porCx, 1#, cx, sn, so) Then lblTempo.Caption = APS_FmtHMS(sn)
End Sub

Private Sub tR1_Change()
    CalcTempo
End Sub

Private Sub tR2_Change()
    CalcTempo
End Sub

Private Sub tR3_Change()
    CalcTempo
End Sub

Private Sub cU1_Change()
    CalcTempo
End Sub

'----------------------------------------------------------
' Carregar a selecao (sempre do cadastro SALVO)
'----------------------------------------------------------
Private Sub CarregarSelecao(ByVal modo As Long, ByVal maq As String, ByVal cod As String)
    Dim v As String, u As String, k As Long, txt As String

    Descarregar

    mBloq = mBloq + 1
    mModo = modo
    mLMaq = maq
    mLCod = cod

    MostrarPainel True

    Select Case modo
        Case 1
            lblGrp1.Caption = "Par" & ChrW(226) & "metros de produ" & ChrW(231) & ChrW(227) & "o"
            Linha 1, "Lote padr" & ChrW(227) & "o (comprimidos)", 136, 120
            Linha 2, "Comprimidos por caixa", 162, 120
            If APS_RegBuscar(PAR_LOTE, maq, cod, v, u) Then txtR(1).Text = APS_ParFmtNum(v)
            If APS_RegBuscar(PAR_CPC, maq, cod, v, u) Then txtR(2).Text = APS_ParFmtNum(v)

            If APS_ParUsaVelocidade(maq) Then
                Linha 3, "Velocidade", 188, 120
                cU1.Clear
                cU1.AddItem UN_CX_MIN
                cU1.AddItem UN_COMP_MIN
                cU1.Width = 130
                UnidadeNaLinha 188
                If APS_RegBuscar(PAR_VEL, maq, cod, v, u) Then
                    txtR(3).Text = APS_ParFmtNum(v)
                    SelecionarItem cU1, u
                End If
                lblTempoL.Top = 217
                lblTempo.Top = 217
                lblTempoL.Visible = True
                lblTempo.Visible = True
                lblHint.Caption = "Tempo de refer" & ChrW(234) & "ncia = lote " & ChrW(247) & " comprimidos por caixa " & ChrW(247) & " velocidade (sem OEE, setup ou paradas)."
                lblHint.Top = 238
            Else
                lblHint.Caption = "Blisterflex: sem velocidade; a dura" & ChrW(231) & ChrW(227) & "o " & ChrW(233) & " informada manualmente na opera" & ChrW(231) & ChrW(227) & "o."
                lblHint.Top = 192
            End If

        Case 2
            lblGrp1.Caption = "Identifica" & ChrW(231) & ChrW(227) & "o e quantidade"
            lblGrp2.Caption = "Processos de fabrica" & ChrW(231) & ChrW(227) & "o"
            lblGrp2.Visible = True
            Linha 1, "Quantidade de comprimidos", 136, 120
            Linha 2, "Peso", 162, 70
            Linha 3, "Granula" & ChrW(231) & ChrW(227) & "o", 218, 220
            Linha 4, "Sala de mistura", 244, 140
            cU1.Clear
            cU1.AddItem "mg"
            cU1.AddItem "g"
            cU1.AddItem "kg"
            cU1.Width = 60
            cU1.Left = 272
            UnidadeNaLinha 162
            cU1.Left = 272
            If APS_RegBuscar(PAR_QTDCOMP, "", cod, v, u) Then txtR(1).Text = APS_ParFmtNum(v)
            If APS_RegBuscar(PAR_PESO, "", cod, v, u) Then
                txtR(2).Text = APS_ParFmtNum(v)
                SelecionarItem cU1, u
            End If
            If APS_RegBuscar(PAR_GRANUL, "", cod, v, u) Then txtR(3).Text = v
            If APS_RegBuscar(PAR_SALA, "", cod, v, u) Then txtR(4).Text = v
            lblHint.Caption = "Informativos: n" & ChrW(227) & "o entram no c" & ChrW(225) & "lculo da dura" & ChrW(231) & ChrW(227) & "o."
            lblHint.Top = 270

        Case 3
            lblGrp1.Caption = "Capacidade da etapa (geral)"
            Linha 1, "Capacidade (kg)", 136, 100
            If APS_RegBuscar(PAR_CAPAC, maq, "", v, u) Then txtR(1).Text = APS_ParFmtNum(v)
            lblHint.Caption = "A dura" & ChrW(231) & ChrW(227) & "o desta etapa " & ChrW(233) & " informada manualmente na opera" & ChrW(231) & ChrW(227) & "o."
            lblHint.Top = 162

            If APS_EhSala(maq) And Len(cod) > 0 Then
                Linha 2, "Meia dosagem do produto (kg)", 162, 100
                If APS_RegBuscar(PAR_MEIA, "Sala de Mistura", cod, v, u) Then txtR(2).Text = APS_ParFmtNum(v)
                lblHint.Caption = "Meia dosagem: vazio = o produto n" & ChrW(227) & "o usa. A capacidade vale para todos os produtos."
                lblHint.Top = 190
            End If
    End Select

    TirarFoto
    CalcTempo

    txt = "  "
    Select Case modo
        Case 1: txt = txt & maq & "   " & ChrW(8250) & "   " & NomeProd(cod)
        Case 2: txt = txt & "Dados do produto   " & ChrW(8250) & "   " & NomeProd(cod)
        Case 3
            txt = txt & maq
            If Len(cod) > 0 Then txt = txt & "   " & ChrW(8250) & "   " & NomeProd(cod)
    End Select

    lblSel.ForeColor = RGB(31, 73, 125)
    lblSel.Caption = txt

    mBloq = mBloq - 1
End Sub

Private Sub cmdCarregar_Click()
    Dim modo As Long, maq As String, cod As String

    lstRes.Visible = False

    modo = ModoSel()
    If modo = 0 Then Msg "Escolha a " & ChrW(225) & "rea.", True: Exit Sub

    If modo = 1 Or modo = 3 Then
        maq = Trim$(cboMaq.Text)
        If Len(maq) = 0 Then
            If modo = 1 Then
                Msg "Escolha a m" & ChrW(225) & "quina.", True
            Else
                Msg "Escolha a etapa.", True
            End If
            Exit Sub
        End If
    End If

    cod = ""
    If modo = 1 Or modo = 2 Then
        If Len(mProdPend) = 0 Then Msg "Pesquise e escolha o produto na lista.", True: Exit Sub
        cod = mProdPend
    ElseIf APS_EhSala(maq) Then
        cod = mProdPend
    End If

    If mModo > 0 And TemMudancas() Then
        If MsgBox("H" & ChrW(225) & " altera" & ChrW(231) & ChrW(245) & "es n" & ChrW(227) & "o salvas. Recarregar e descart" & ChrW(225) & "-las?", vbYesNo + vbQuestion, APS_TITULO) <> vbYes Then Exit Sub
    End If

    CarregarSelecao modo, maq, cod
    Msg "", False
End Sub

'----------------------------------------------------------
' Salvar / Desfazer / Fechar
'----------------------------------------------------------
Private Sub AddIt(ByRef n As Long, ByRef t() As Long, ByRef m() As String, ByRef c() As String, _
                  ByRef v() As String, ByRef u() As String, _
                  ByVal tp As Long, ByVal mq As String, ByVal cd As String, ByVal vl As String, ByVal un As String)
    n = n + 1
    t(n) = tp
    m(n) = mq
    c(n) = cd
    v(n) = vl
    u(n) = un
End Sub

' Itens do que esta na tela (somente a selecao atual)
Private Function Montar(ByRef t() As Long, ByRef m() As String, ByRef c() As String, ByRef v() As String, ByRef u() As String) As Long
    Dim n As Long

    ReDim t(1 To 6)
    ReDim m(1 To 6)
    ReDim c(1 To 6)
    ReDim v(1 To 6)
    ReDim u(1 To 6)

    Select Case mModo
        Case 1
            AddIt n, t, m, c, v, u, PAR_LOTE, mLMaq, mLCod, Trim$(txtR(1).Text), "Comprimidos"
            AddIt n, t, m, c, v, u, PAR_CPC, mLMaq, mLCod, Trim$(txtR(2).Text), "Comprimidos/caixa"
            If APS_ParUsaVelocidade(mLMaq) Then AddIt n, t, m, c, v, u, PAR_VEL, mLMaq, mLCod, Trim$(txtR(3).Text), Trim$(cU1.Text)
        Case 2
            AddIt n, t, m, c, v, u, PAR_QTDCOMP, "", mLCod, Trim$(txtR(1).Text), "Comprimidos"
            AddIt n, t, m, c, v, u, PAR_PESO, "", mLCod, Trim$(txtR(2).Text), Trim$(cU1.Text)
            AddIt n, t, m, c, v, u, PAR_GRANUL, "", mLCod, Trim$(txtR(3).Text), "Texto"
            AddIt n, t, m, c, v, u, PAR_SALA, "", mLCod, Trim$(txtR(4).Text), "Sala"
        Case 3
            AddIt n, t, m, c, v, u, PAR_CAPAC, mLMaq, "", Trim$(txtR(1).Text), "kg"
            If APS_EhSala(mLMaq) And Len(mLCod) > 0 Then AddIt n, t, m, c, v, u, PAR_MEIA, "Sala de Mistura", mLCod, Trim$(txtR(2).Text), "kg"
    End Select

    Montar = n
End Function

Private Sub cmdSalvar_Click()
    Dim t() As Long, m() As String, c() As String, v() As String, u() As String
    Dim n As Long, i As Long, ng As Long, nr As Long, txtErr As String
    Dim ev As String, eu As String, lstRem As String

    If mModo = 0 Then Exit Sub
    If Not TemMudancas() Then Msg "Nenhuma altera" & ChrW(231) & ChrW(227) & "o para salvar.", False: Exit Sub

    n = Montar(t, m, c, v, u)

    ' campos vazios que hoje existem no cadastro serao removidos: pedir confirmacao
    For i = 1 To n
        If Len(v(i)) = 0 Then
            If APS_RegBuscar(t(i), m(i), c(i), ev, eu) Then
                lstRem = lstRem & vbCrLf & " - " & APS_RegNomeTipo(t(i))
            End If
        End If
    Next i

    If Len(lstRem) > 0 Then
        If MsgBox("Os campos abaixo est" & ChrW(227) & "o vazios e ser" & ChrW(227) & "o REMOVIDOS do cadastro:" & vbCrLf & lstRem & vbCrLf & vbCrLf & _
                  "Opera" & ChrW(231) & ChrW(245) & "es j" & ChrW(225) & " gravadas n" & ChrW(227) & "o s" & ChrW(227) & "o alteradas. Continuar?", vbYesNo + vbQuestion, APS_TITULO) <> vbYes Then
            Msg "Nada foi salvo.", True
            Exit Sub
        End If
    End If

    If Not APS_RegAplicar(n, t, m, c, v, u, txtErr, ng, nr) Then
        Msg txtErr, True
        Exit Sub
    End If

    CarregarSelecao mModo, mLMaq, mLCod

    Msg "Salvo e conferido no cadastro: " & ng & " gravado(s), " & nr & " removido(s)." & vbCrLf & _
        "Vale para opera" & ChrW(231) & ChrW(245) & "es novas ou editadas; opera" & ChrW(231) & ChrW(245) & "es j" & ChrW(225) & " gravadas n" & ChrW(227) & "o foram alteradas. " & _
        "Para manter ap" & ChrW(243) & "s fechar o Excel, salve a planilha (Ctrl+S).", False
End Sub

Private Sub cmdDesf_Click()
    If mModo = 0 Then Exit Sub
    CarregarSelecao mModo, mLMaq, mLCod
    Msg "Valores recarregados do cadastro salvo.", False
End Sub

Private Sub cmdFechar_Click()
    If TemMudancas() Then
        If MsgBox("H" & ChrW(225) & " altera" & ChrW(231) & ChrW(245) & "es n" & ChrW(227) & "o salvas. Fechar mesmo assim?", vbYesNo + vbQuestion, APS_TITULO) <> vbYes Then Exit Sub
    End If
    Me.Hide
End Sub

Private Sub UserForm_Click()
    lstRes.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        cmdFechar_Click
    End If
End Sub
