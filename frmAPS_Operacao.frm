VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAPS_Operacao 
   Caption         =   "Operacao"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAPS_Operacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

'==========================================================
' APS PURAN - Formulario Nova / Editar Operacao
' Controles criados em codigo
'
' Nova / Editar / Excluir Operacao
'
' Codigo do produto > Produto > Lote > Maquina > Data >
' Horario > Quantidade
'
' Mediseal: duracao automatica
' Demais maquinas: duracao manual
'==========================================================

Private WithEvents cboCodigo As MSForms.ComboBox
Attribute cboCodigo.VB_VarHelpID = -1
Private WithEvents cboLote As MSForms.ComboBox
Attribute cboLote.VB_VarHelpID = -1
Private WithEvents cboMaquina As MSForms.ComboBox
Attribute cboMaquina.VB_VarHelpID = -1
Private WithEvents txtQtd As MSForms.TextBox
Attribute txtQtd.VB_VarHelpID = -1
Private WithEvents txtData As MSForms.TextBox
Attribute txtData.VB_VarHelpID = -1
Private WithEvents txtInicio As MSForms.TextBox
Attribute txtInicio.VB_VarHelpID = -1
Private WithEvents txtOEE As MSForms.TextBox
Attribute txtOEE.VB_VarHelpID = -1
Private WithEvents txtDur As MSForms.TextBox
Attribute txtDur.VB_VarHelpID = -1
Private WithEvents cmdSalvar As MSForms.CommandButton
Attribute cmdSalvar.VB_VarHelpID = -1
Private WithEvents cmdCancelar As MSForms.CommandButton
Attribute cmdCancelar.VB_VarHelpID = -1
Private WithEvents cmdExcluir As MSForms.CommandButton
Attribute cmdExcluir.VB_VarHelpID = -1
Private txtProduto As MSForms.TextBox
Private lblLoteFinal As MSForms.Label
Private lblCad As MSForms.Label
Private lblTipo As MSForms.Label
Private cboStatus As MSForms.ComboBox
Private txtObs As MSForms.TextBox
Private lblCalc As MSForms.Label
Private lblAviso As MSForms.Label

Private WithEvents optSetupSim As MSForms.OptionButton
Attribute optSetupSim.VB_VarHelpID = -1
Private WithEvents optSetupNao As MSForms.OptionButton
Attribute optSetupNao.VB_VarHelpID = -1
Private WithEvents cboSetupParam As MSForms.ComboBox
Attribute cboSetupParam.VB_VarHelpID = -1
Private lblSetupInfo As MSForms.Label

Private mBloq As Long
Private mRecalc As Boolean

' Atraso: tempo ADICIONAL de programacao no fim da producao (nao altera a quantidade nem a duracao original)
Private WithEvents txtAtraso As MSForms.TextBox
Attribute txtAtraso.VB_VarHelpID = -1
Private mNovo As Boolean
Private mID As String
Private mLinha As Long
Private mDurLegada As Double

' OP (Ordem de Producao): campo proprio, separado do ID tecnico (mID, interno e oculto)
Private txtOP As MSForms.TextBox
' valores GRAVADOS na operacao (edicao): velocidade-base, produto de referencia e OEE inferido
Private mVelGravada As Double
Private mCodGravado As String
Private mOEEInferido As Boolean
Private mOps() As tOperacao
Private mN As Long
Private mTinhaSetupAoAbrir As Boolean
Private mSetupRotuloExistente As String

' Parametros (Cadastros e Parametros): lote padrao preenchido automaticamente / FETTE antiga mantem duracao manual
Private mAutoQtd As String          ' ultimo valor preenchido AUTOMATICAMENTE no campo Quantidade (nunca sobrescreve digitacao)
Private mLegadoManual As Boolean    ' edicao de operacao FETTE antiga (sem quantidade em comprimidos): duracao segue manual
Private mMaqGravada As String       ' maquina gravada da operacao em edicao

'----------------------------------------------------------
' Montagem da tela
'----------------------------------------------------------
Private Sub UserForm_Initialize()
    MontarTela
End Sub

Private Function NovoLabel(ByVal nome As String, ByVal txt As String, ByVal x As Single, _
                           ByVal y As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1", nome, True)
    l.Left = x
    l.Top = y
    l.Width = w
    l.Height = h
    l.Caption = txt
    l.Font.Size = 9
    Set NovoLabel = l
End Function

Private Function NovoCombo(ByVal nome As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal estilo As Long) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    c.Left = x
    c.Top = y
    c.Width = w
    c.Height = 18
    c.Style = estilo
    c.Font.Size = 9
    If estilo = 0 Then c.MatchEntry = 2
    Set NovoCombo = c
End Function

Private Function NovoTexto(ByVal nome As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = Me.Controls.Add("Forms.TextBox.1", nome, True)
    t.Left = x
    t.Top = y
    t.Width = w
    t.Height = h
    t.Font.Size = 9
    Set NovoTexto = t
End Function

Private Sub MontarTela()
    Dim y As Single, lx As Single, cx As Single, cw As Single, p As Single
    Dim l As MSForms.Label

    Me.Width = 690
    Me.Height = 535
    Me.StartUpPosition = 1

    lx = 12
    cx = 132
    cw = 200
    p = 25
    y = 12

    NovoLabel "lb1", "C" & ChrW(243) & "digo do Produto", lx, y + 3, 118, 14
    Set cboCodigo = NovoCombo("cboCodigo", cx, y, cw, 0)
    y = y + p

    NovoLabel "lb2", "Produto", lx, y + 3, 118, 14
    Set txtProduto = NovoTexto("txtProduto", cx, y, cw, 18)
    txtProduto.Locked = True
    txtProduto.BackColor = RGB(238, 238, 238)
    y = y + p - 4

    Set lblLoteFinal = NovoLabel("lblLoteFinal", "", cx, y, cw, 12)
    lblLoteFinal.Font.Size = 8
    lblLoteFinal.ForeColor = RGB(90, 90, 90)
    y = y + 16

    NovoLabel "lbOP", "OP", lx, y + 3, 118, 14
    Set txtOP = NovoTexto("txtOP", cx, y, cw, 18)
    txtOP.MaxLength = 30
    y = y + p

    NovoLabel "lb3", "Lote", lx, y + 3, 118, 14
    Set cboLote = NovoCombo("cboLote", cx, y, cw, 0)
    y = y + p

    NovoLabel "lb4", "M" & ChrW(225) & "quina", lx, y + 3, 118, 14
    Set cboMaquina = NovoCombo("cboMaquina", cx, y, cw, 2)
    y = y + p

    NovoLabel "lb5", "Data de in" & ChrW(237) & "cio", lx, y + 3, 118, 14
    Set txtData = NovoTexto("txtData", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb6", "Hor" & ChrW(225) & "rio de in" & ChrW(237) & "cio", lx, y + 3, 118, 14
    Set txtInicio = NovoTexto("txtInicio", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb7", "Quantidade (comprimidos)", lx, y + 3, 118, 14
    Set txtQtd = NovoTexto("txtQtd", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb8", "Tipo de dura" & ChrW(231) & ChrW(227) & "o", lx, y + 3, 118, 14
    Set lblTipo = NovoLabel("lblTipo", "", cx, y + 3, cw, 14)
    lblTipo.Font.Bold = True
    y = y + p

    NovoLabel "lb9", "OEE (%)", lx, y + 3, 118, 14
    Set txtOEE = NovoTexto("txtOEE", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb10", "Dura" & ChrW(231) & ChrW(227) & "o (hh:mm)", lx, y + 3, 118, 14
    Set txtDur = NovoTexto("txtDur", cx, y, cw, 18)
    y = y + p

    NovoLabel "lb11", "Planejamento", lx, y + 3, 118, 14
    Set cboStatus = NovoCombo("cboStatus", cx, y, cw, 2)
    y = y + p

    NovoLabel "lb12", "Observa" & ChrW(231) & ChrW(245) & "es", lx, y + 3, 118, 14
    Set txtObs = NovoTexto("txtObs", cx, y, cw, 40)
    txtObs.Multiline = True
    txtObs.EnterKeyBehavior = True
    txtObs.ScrollBars = 2

    y = y + 40 + 8
    NovoLabel "lbSetup1", "Prepara" & ChrW(231) & ChrW(227) & "o", lx, y + 3, 118, 14
    Set optSetupSim = Me.Controls.Add("Forms.OptionButton.1", "optSetupSim", True)
    optSetupSim.Left = cx: optSetupSim.Top = y: optSetupSim.Width = 110: optSetupSim.Height = 18
    optSetupSim.Caption = "Setup + Limpeza": optSetupSim.GroupName = "grpSetup": optSetupSim.Font.Size = 8
    Set optSetupNao = Me.Controls.Add("Forms.OptionButton.1", "optSetupNao", True)
    optSetupNao.Left = cx + 112: optSetupNao.Top = y: optSetupNao.Width = 70: optSetupNao.Height = 18
    optSetupNao.Caption = "Nenhuma": optSetupNao.GroupName = "grpSetup": optSetupNao.Font.Size = 8
    optSetupNao.Value = True
    y = y + 22

    NovoLabel "lbSetup2", "Par" & ChrW(226) & "metro", lx, y + 3, 118, 14
    Set cboSetupParam = NovoCombo("cboSetupParam", cx, y, cw, 2)
    y = y + 22

    Set lblSetupInfo = NovoLabel("lblSetupInfo", "", cx, y, cw, 28)
    lblSetupInfo.Font.Size = 8
    lblSetupInfo.ForeColor = RGB(140, 100, 0)
    lblSetupInfo.WordWrap = True

    Set l = NovoLabel("lbTitCalc", "C" & ChrW(225) & "lculo autom" & ChrW(225) & "tico", 360, 12, 310, 16)
    l.Font.Bold = True

    Set lblCalc = NovoLabel("lblCalc", "", 360, 32, 310, 210)
    lblCalc.WordWrap = True
    lblCalc.BorderStyle = 1
    lblCalc.BackColor = RGB(248, 250, 252)

    Set lblAviso = NovoLabel("lblAviso", "", 360, 250, 310, 76)
    lblAviso.WordWrap = True

    ' resumo do Cadastro de Parametros de Produto (produto + maquina escolhidos): mostra o que existe e o que falta
    Set lblCad = NovoLabel("lblCad", "", 360, 330, 310, 46)
    lblCad.WordWrap = True
    lblCad.Font.Size = 8
    lblCad.ForeColor = RGB(90, 90, 90)

    '------------------------------------------------------
    ' Botoes
    '------------------------------------------------------
    Set cmdSalvar = Me.Controls.Add("Forms.CommandButton.1", "cmdSalvar", True)
    cmdSalvar.Caption = "Salvar"
    cmdSalvar.Left = 350
    cmdSalvar.Top = 455
    cmdSalvar.Width = 90
    cmdSalvar.Height = 28
    cmdSalvar.Font.Bold = True

    Set cmdCancelar = Me.Controls.Add("Forms.CommandButton.1", "cmdCancelar", True)
    cmdCancelar.Caption = "Cancelar"
    cmdCancelar.Left = 455
    cmdCancelar.Top = 455
    cmdCancelar.Width = 90
    cmdCancelar.Height = 28

    Set cmdExcluir = Me.Controls.Add("Forms.CommandButton.1", "cmdExcluir", True)
    cmdExcluir.Caption = "Excluir Opera" & ChrW(231) & ChrW(227) & "o"
    cmdExcluir.Left = 560
    cmdExcluir.Top = 455
    cmdExcluir.Width = 110
    cmdExcluir.Height = 28
    cmdExcluir.Font.Bold = True
    cmdExcluir.ForeColor = RGB(192, 0, 0)

    ' Atraso visual: estende SO o card para a direita (nao muda inicio, fim, duracao nem quantidade)
    NovoLabel "lbAtr", "Atraso (hh:mm)", 360, 378, 100, 14
    Set txtAtraso = NovoTexto("txtAtraso", 465, 375, 80, 18)
    Set l = NovoLabel("lbAtrInfo", "Tempo adicional no final da produ" & ChrW(231) & ChrW(227) & "o: ocupa a m" & ChrW(225) & "quina e move as seguintes em cascata. N" & ChrW(227) & "o altera a quantidade nem a dura" & ChrW(231) & ChrW(227) & "o original. Vazio = sem atraso.", 360, 396, 310, 30)
    l.Font.Size = 8
    l.ForeColor = RGB(90, 90, 90)
    l.WordWrap = True
End Sub

'----------------------------------------------------------
' Carga (nova / editar)
'----------------------------------------------------------
Public Sub Iniciar(ByVal id As String)
    Dim i As Long, k As Long, st As Variant
    Dim cods() As String, nomes() As String, lotes() As Double, vels() As Double
    Dim p1 As Double, p2 As Double

    mBloq = mBloq + 1
    On Error GoTo Falha

    mNovo = (Len(id) = 0)
    mAutoQtd = ""
    mLegadoManual = False
    mMaqGravada = ""
    txtAtraso.Text = ""
    mN = APS_LerOps(mOps)
    If mN < 0 Then mN = 0

    k = APS_Produtos(cods, nomes, lotes, vels)

    cboCodigo.Clear
    For i = 1 To k
        cboCodigo.AddItem cods(i)
    Next i

    cboStatus.Clear
    st = APS_ListaStatus()

    For i = LBound(st) To UBound(st)
        cboStatus.AddItem CStr(st(i))
    Next i

    If mNovo Then
        mID = APS_ProximoID()
        mLinha = 0

        Me.Caption = "Nova Opera" & ChrW(231) & ChrW(227) & "o"

        SelecionarItem cboStatus, APS_ST_PLANEJADA, False

        cmdExcluir.Visible = False

        If APS_DiasVisiveis(p1, p2) Then
            If Int(Date) >= p1 And Int(Date) <= p2 Then
                txtData.Text = Format$(Date, "dd/mm/yyyy")
            Else
                txtData.Text = Format$(p1, "dd/mm/yyyy")
            End If
        Else
            txtData.Text = Format$(Date, "dd/mm/yyyy")
        End If

        txtOEE.Text = APS_FmtQtd(APS_OEEPadrao() * 100#)

    Else
        mID = id

        Me.Caption = "Editar Opera" & ChrW(231) & ChrW(227) & "o"

        cmdExcluir.Visible = True

        CarregarExistente
    End If

Sai:
    mBloq = mBloq - 1

    AtualizarProduto
    AtualizarHabilitacao

    If mTinhaSetupAoAbrir Then
        mBloq = mBloq + 1
        optSetupSim.Value = True
        SelecionarItem cboSetupParam, mSetupRotuloExistente, True
        mBloq = mBloq - 1
        AtualizarSetupInfo
    End If

    Recalcular

    Exit Sub

Falha:
    MsgBox "Erro ao abrir a opera" & ChrW(231) & ChrW(227) & "o: " & Err.Description, vbCritical, APS_TITULO
    Resume Sai
End Sub

Private Sub CarregarExistente()
    Dim i As Long, k As Long, o As tOperacao
    Dim vb As Double, inf As Double, atrIni As Long

    For i = 1 To mN
        If APS_Ig(mOps(i).id, mID) Then
            k = i
            Exit For
        End If
    Next i

    If k = 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o n" & ChrW(227) & "o encontrada.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    o = mOps(k)
    mLinha = o.linha

    txtOP.Text = o.ordem
    If Len(Trim$(o.ordem)) > 0 Then
        Me.Caption = "Editar Opera" & ChrW(231) & ChrW(227) & "o " & ChrW(8212) & " OP " & Trim$(o.ordem)
    End If

    mCodGravado = o.codigo
    mVelGravada = o.velBase

    cboCodigo.Text = o.codigo

    AtualizarProduto
    CarregarLotes

    cboLote.Text = o.lote

    CarregarMaquinas
    SelecionarItem cboMaquina, o.maquina, True

    If o.comp > 0 Then
        txtQtd.Text = APS_FmtQtd(o.comp)
    ElseIf o.caixas > 0 Then
        ' operacao antiga (so caixas): mostra os comprimidos equivalentes (caixas x 30, a mesma conversao ja usada pela Mediseal).
        ' Nada e regravado ate o usuario salvar.
        txtQtd.Text = APS_FmtQtd(Round(o.caixas * APS_CP_CAIXA, 0))
    End If

    mMaqGravada = o.maquina
    mLegadoManual = APS_EhFette(o.maquina) And Not (o.comp > 0) And Not o.SemDur

    If Not o.SemDur Then
        mDurLegada = APS_DurOriginal(o)      ' duracao ORIGINAL (sem o atraso)
    End If

    txtData.Text = Format$(Int(o.ini + APS_EPS), "dd/mm/yyyy")
    txtInicio.Text = Format$(o.ini - Int(o.ini + APS_EPS), "hh:nn")

    mOEEInferido = False
    If o.oee > 0 Then
        txtOEE.Text = APS_FmtQtd(o.oee * 100#)
    Else
        txtOEE.Text = APS_FmtQtd(APS_OEEPadrao() * 100#)
        ' Operacao SEM OEE gravado (anterior a esta versao): para a edicao nao trocar o OEE em silencio pelo
        ' padrao, ele e inferido da duracao e da quantidade que estao gravadas (so se ficar entre 1% e 100%).
        If APS_EhMediseal(o.maquina) And o.caixas > 0 And APS_DurOriginal(o) > 0 And Not o.SemDur Then
            If o.velBase > 0 Then vb = o.velBase Else vb = APS_VelBaseDoCodigo(Trim$(o.codigo))
            If vb > 0 Then
                inf = (o.caixas * APS_CP_CAIXA) / (vb * APS_DurOriginal(o) * 1440#)
                If inf >= 0.01 And inf <= 1.0005 Then
                    If inf > 1# Then inf = 1#
                    txtOEE.Text = APS_FmtQtd(Round(inf * 100#, 1))
                    mOEEInferido = True
                End If
            End If
        End If
    End If

    If Not APS_EhMediseal(o.maquina) And Not o.SemDur Then
        txtDur.Text = Format$(Int(APS_DurOriginal(o) * 24# + APS_EPS), "0") & ":" & _
                      Format$(Round((APS_DurOriginal(o) * 24# - Int(APS_DurOriginal(o) * 24# + APS_EPS)) * 60#, 0), "00")
    End If

    SelecionarItem cboStatus, o.status, True

    txtObs.Text = o.obs

    atrIni = CLng(Round(o.atraso * 1440#, 0))
    If atrIni > 0 Then txtAtraso.Text = Format$(atrIni \ 60, "0") & ":" & Format$(atrIni Mod 60, "00")

    Dim idxSetup As Long
    idxSetup = APS_AcharSetupDaProducao(mOps, mN, mID)
    If idxSetup > 0 Then
        mTinhaSetupAoAbrir = True
        mSetupRotuloExistente = mOps(idxSetup).rotAux
        ' com preparacao, o campo de horario e o INICIO DA PREPARACAO (inicio do bloco), nao o da producao
        txtData.Text = Format$(Int(mOps(idxSetup).ini + APS_EPS), "dd/mm/yyyy")
        txtInicio.Text = Format$(mOps(idxSetup).ini - Int(mOps(idxSetup).ini + APS_EPS), "hh:nn")
    Else
        mTinhaSetupAoAbrir = False
        mSetupRotuloExistente = ""
    End If
End Sub

Private Sub SelecionarItem(ByVal cbo As MSForms.ComboBox, ByVal txt As String, ByVal adicionar As Boolean)
    Dim i As Long

    For i = 0 To cbo.ListCount - 1
        If APS_Ig(CStr(cbo.List(i)), txt) Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    If adicionar And Len(txt) > 0 Then
        cbo.AddItem txt
        cbo.ListIndex = cbo.ListCount - 1
    End If
End Sub

'----------------------------------------------------------
' Cascata
'----------------------------------------------------------
Private Sub AtualizarProduto()
    Dim i As Long
    Dim cods() As String, nomes() As String
    Dim lotes() As Double, vels() As Double
    Dim k As Long

    txtProduto.Text = ""
    lblLoteFinal.Caption = ""

    k = APS_Produtos(cods, nomes, lotes, vels)

    For i = 1 To k
        If cods(i) = Trim$(cboCodigo.Text) Then
            txtProduto.Text = nomes(i)

            If lotes(i) > 0 Then
                lblLoteFinal.Caption = "Lote final: " & APS_FmtQtd(lotes(i)) & " caixas"
            End If

            Exit For
        End If
    Next i
End Sub

Private Sub CarregarLotes()
    Dim lst() As String, n As Long, i As Long

    cboLote.Clear

    If Len(txtProduto.Text) = 0 Then Exit Sub

    n = APS_LotesDoProduto(mOps, mN, Trim$(cboCodigo.Text), lst)

    For i = 1 To n
        cboLote.AddItem lst(i)
    Next i
End Sub

Private Sub CarregarMaquinas()
    Dim lst() As String, n As Long, i As Long

    cboMaquina.Clear

    If Len(txtProduto.Text) = 0 Then Exit Sub

    n = APS_MaquinasDoProduto(Trim$(cboCodigo.Text), lst)

    For i = 1 To n
        cboMaquina.AddItem lst(i)
    Next i
End Sub

' Duracao calculada pelo parametro? Mediseal sempre; FETTE so com velocidade cadastrada (e nao em operacao FETTE antiga).
Private Function DurAuto() As Boolean
    If Len(cboMaquina.Text) = 0 Then Exit Function
    If mLegadoManual And APS_Ig(cboMaquina.Text, mMaqGravada) Then Exit Function
    DurAuto = APS_DurAutomatica(cboMaquina.Text, Trim$(cboCodigo.Text))
End Function

' Lote padrao (comprimidos) da combinacao maquina + produto. So preenche com o campo vazio ou ainda com o ultimo valor
' automatico; o que o usuario digitou nunca e sobrescrito, e a edicao NUNCA grava no cadastro de parametros.
Private Sub PreencherLotePadrao()
    Dim lp As Double, cur As String

    If Len(cboMaquina.Text) = 0 Or Len(Trim$(cboCodigo.Text)) = 0 Then Exit Sub

    lp = APS_ParLote(cboMaquina.Text, Trim$(cboCodigo.Text))
    cur = Trim$(txtQtd.Text)

    If Len(cur) > 0 And cur <> mAutoQtd Then Exit Sub

    mBloq = mBloq + 1

    If lp > 0 Then
        txtQtd.Text = APS_FmtQtd(lp)
        mAutoQtd = txtQtd.Text
    ElseIf Len(mAutoQtd) > 0 And cur = mAutoQtd Then
        txtQtd.Text = ""
        mAutoQtd = ""
    End If

    mBloq = mBloq - 1
End Sub

Private Sub AtualizarHabilitacao()
    Dim tProd As Boolean
    Dim tLote As Boolean
    Dim tMaq As Boolean
    Dim med As Boolean

    tProd = (Len(txtProduto.Text) > 0)
    tLote = tProd And (Len(Trim$(cboLote.Text)) > 0)
    tMaq = tLote And (Len(cboMaquina.Text) > 0)
    med = tMaq And DurAuto()

    cboLote.Enabled = tProd
    cboMaquina.Enabled = tLote

    txtData.Enabled = tMaq
    txtInicio.Enabled = tMaq
    txtQtd.Enabled = tMaq

    txtOEE.Enabled = med

    txtDur.Enabled = tMaq And Not med
    txtDur.Locked = med

    cmdSalvar.Enabled = tMaq

    If Not tMaq Then
        lblTipo.Caption = ""
    ElseIf med Then
        lblTipo.Caption = "Autom" & ChrW(225) & "tica (" & cboMaquina.Text & ")"
    Else
        lblTipo.Caption = "Manual"
    End If

    AtualizarSetupCombo
End Sub

'----------------------------------------------------------
' PREPARACAO (Setup + Limpeza, opcional, escolhida pelo usuario)
'----------------------------------------------------------
Private Sub AtualizarSetupCombo()
    Dim rot() As String, minu() As Long, k As Long, i As Long, atual As String

    If cboSetupParam Is Nothing Then Exit Sub

    mBloq = mBloq + 1          ' limpar/preencher a lista nao deve disparar a previa a cada item
    atual = cboSetupParam.Text
    cboSetupParam.Clear

    If Len(cboMaquina.Text) > 0 Then
        k = APS_ParametrosSetup(cboMaquina.Text, rot, minu)
        For i = 1 To k
            cboSetupParam.AddItem rot(i)
        Next i
    End If

    SelecionarItem cboSetupParam, atual, False
    mBloq = mBloq - 1

    AtualizarSetupInfo
End Sub

Private Sub AtualizarSetupInfo()
    Dim tMaq As Boolean

    If optSetupSim Is Nothing Or optSetupNao Is Nothing Or cboSetupParam Is Nothing Or lblSetupInfo Is Nothing Then Exit Sub

    tMaq = (Len(cboMaquina.Text) > 0)

    optSetupSim.Enabled = tMaq
    optSetupNao.Enabled = tMaq
    cboSetupParam.Enabled = tMaq And optSetupSim.Value And cboSetupParam.ListCount > 0

    If Not tMaq Then
        lblSetupInfo.Caption = ""
    ElseIf cboSetupParam.ListCount = 0 Then
        lblSetupInfo.Caption = "N" & ChrW(227) & "o h" & ChrW(225) & " par" & ChrW(226) & "metro de prepara" & ChrW(231) & ChrW(227) & "o cadastrado para esta m" & ChrW(225) & "quina."
    ElseIf optSetupSim.Value Then
        lblSetupInfo.Caption = "O hor" & ChrW(225) & "rio acima passa a ser o IN" & ChrW(205) & "CIO DA PREPARA" & ChrW(199) & ChrW(195) & "O. A produ" & ChrW(231) & ChrW(227) & "o come" & ChrW(231) & "a quando Setup + Limpeza terminam (pelo calend" & ChrW(225) & "rio)."
    Else
        lblSetupInfo.Caption = ""
    End If
End Sub

Private Sub optSetupSim_Click()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
    Recalcular
End Sub

Private Sub optSetupNao_Click()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
    Recalcular
End Sub

Private Sub cboSetupParam_Change()
    If mBloq > 0 Then Exit Sub
    AtualizarSetupInfo
    Recalcular
End Sub

Private Sub cboCodigo_Change()
    Dim anterior As String

    If mBloq > 0 Then Exit Sub

    anterior = txtProduto.Text

    mBloq = mBloq + 1

    AtualizarProduto

    If txtProduto.Text <> anterior Then
        cboLote.Clear
        cboLote.Text = ""

        cboMaquina.Clear

        txtDur.Text = ""

        CarregarLotes
    End If

    mBloq = mBloq - 1

    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub cboLote_Change()
    If mBloq > 0 Then Exit Sub

    mBloq = mBloq + 1

    If Len(cboMaquina.Text) = 0 And Len(txtProduto.Text) > 0 Then
        CarregarMaquinas
    End If

    mBloq = mBloq - 1

    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub cboMaquina_Change()
    If mBloq > 0 Then Exit Sub

    PreencherLotePadrao
    AtualizarHabilitacao
    Recalcular
End Sub

Private Sub txtQtd_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtData_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtInicio_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtAtraso_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtOEE_Change()
    If mBloq > 0 Then Exit Sub
    mOEEInferido = False
    Recalcular
End Sub

Private Sub txtDur_Change()
    If mBloq > 0 Then Exit Sub
    Recalcular
End Sub

Private Sub txtQtd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8, 44, 46
        Case Else
            KeyAscii = 0
    End Select
End Sub

'----------------------------------------------------------
' Duracao e calculo
'----------------------------------------------------------
Private Function ParaOEE() As Double
    Dim s As String, v As Double

    s = Replace(Replace(Trim$(txtOEE.Text), "%", ""), ",", ".")

    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9.]*" Then Exit Function

    v = Val(s)

    If v <= 0 Or v > 100 Then Exit Function

    ParaOEE = v / 100#
End Function

Private Function ObterDuracao(ByRef dur As Double, ByRef info As String, _
                              ByRef velBase As Double, ByRef oee As Double) As Boolean
    Dim comp As Double
    Dim compX As Double
    Dim caixas As Double
    Dim cpc As Double
    Dim velNom As Double
    Dim porCaixa As Boolean
    Dim vEf As Double
    Dim segNom As Double
    Dim segOee As Double
    Dim mins As Double
    Dim cod As String
    Dim maq As String
    Dim unid As String
    Dim d As Double

    comp = APS_ParaQtd(txtQtd.Text)
    cod = Trim$(cboCodigo.Text)
    maq = cboMaquina.Text

    If comp > 0 Then
        If comp <> Int(comp) Then
            info = "A quantidade do lote deve ser um n" & ChrW(250) & "mero inteiro de comprimidos."
            Exit Function
        End If
    End If

    If DurAuto() And APS_ParVel(maq, cod) > 0 Then

        ' PARAMETROS CADASTRADOS (FETTE / Mediseal): comprimidos -> caixas -> tempo nominal -> tempo com OEE
        velNom = APS_ParVel(maq, cod)
        unid = APS_ParVelUnid(maq, cod)
        porCaixa = APS_Ig(unid, UN_CX_MIN)
        cpc = APS_ParCPC(maq, cod)
        velBase = 0
        oee = ParaOEE()

        If comp > 0 Then

            If porCaixa And cpc <= 0 Then
                info = "Conte" & ChrW(250) & "do por caixa n" & ChrW(227) & "o cadastrado para " & maq & " / produto " & cod & "." & vbCr & _
                       "Cadastre em Cadastros e Par" & ChrW(226) & "metros."
                Exit Function
            End If

            If oee <= 0 Then
                info = "Informe um OEE v" & ChrW(225) & "lido (1 a 100%)."
                Exit Function
            End If

            If Not APS_CalcTempo(comp, cpc, velNom, porCaixa, oee, caixas, segNom, segOee) Then
                info = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel calcular a dura" & ChrW(231) & ChrW(227) & "o com os par" & ChrW(226) & "metros cadastrados."
                Exit Function
            End If

            dur = APS_ArredMin(segOee / 86400#)
            If dur <= 0 Then dur = 1# / 1440#

            info = "Quantidade: " & APS_FmtQtd(comp) & " comprimidos"

            If cpc > 0 Then
                info = info & vbCr & "Convers" & ChrW(227) & "o: " & APS_FmtQtd(Round(caixas, 2)) & " caixas (" & APS_FmtQtd(cpc) & " comprimidos/caixa)"
            End If

            info = info & vbCr & "Velocidade nominal: " & APS_FmtQtd(velNom) & " " & unid & vbCr & _
                   "Tempo nominal (sem OEE): " & APS_FmtHMS(segNom) & vbCr & _
                   "OEE: " & APS_FmtQtd(oee * 100#) & "%"

            If oee < 0.9999999 Then
                info = info & vbCr & "Tempo com OEE: " & APS_FmtHMS(segOee)
            End If

            info = info & vbCr & "Programado (arredondado ao minuto): " & APS_FmtDur(dur)

            ObterDuracao = True

        ElseIf mDurLegada > 0 Then

            dur = mDurLegada

            info = "Dura" & ChrW(231) & ChrW(227) & "o mantida da opera" & ChrW(231) & ChrW(227) & "o: " & _
                   APS_FmtDur(dur) & vbCr & _
                   "(informe a quantidade em comprimidos para recalcular)"

            ObterDuracao = True

        Else

            info = "Informe a quantidade do lote em comprimidos."

        End If

    ElseIf APS_EhMediseal(maq) Then

        ' Mediseal SEM velocidade cadastrada em Cadastros e Parametros: mantem o calculo ja existente
        ' (velocidade-base em comprimidos/min x OEE). A quantidade e informada em comprimidos.
        velBase = APS_VelBaseDoCodigo(cod)
        ' operacao ja gravada com velocidade-base propria e mesmo produto: usa a GRAVADA, nao o cadastro de hoje
        If mVelGravada > 0 And APS_Ig(cod, mCodGravado) Then velBase = mVelGravada
        oee = ParaOEE()

        If comp > 0 Then

            If velBase <= 0 Then
                info = "Velocidade n" & ChrW(227) & "o cadastrada para este produto na Mediseal." & vbCr & _
                       "Cadastre em Cadastros e Par" & ChrW(226) & "metros."
                Exit Function
            End If

            If oee <= 0 Then
                info = "Informe um OEE v" & ChrW(225) & "lido (1 a 100%)."
                Exit Function
            End If

            caixas = comp / APS_CP_CAIXA

            If Not APS_CalcMediseal(caixas, velBase, oee, compX, vEf, mins) Then Exit Function

            dur = mins / 1440#

            info = "Velocidade-base: " & APS_FmtQtd(velBase) & " comprimidos/min" & vbCr & _
                   "OEE: " & APS_FmtQtd(oee * 100#) & "%" & vbCr & _
                   "Velocidade efetiva: " & APS_FmtQtd(vEf) & " comprimidos/min" & vbCr & _
                   "Quantidade: " & APS_FmtQtd(comp) & " comprimidos" & vbCr & _
                   "Dura" & ChrW(231) & ChrW(227) & "o calculada: " & APS_FmtDur(dur)

            If mOEEInferido Then
                info = info & vbCr & "(OEE inferido da dura" & ChrW(231) & ChrW(227) & "o e da quantidade j" & ChrW(225) & " gravadas)"
            End If

            ObterDuracao = True

        ElseIf mDurLegada > 0 Then

            dur = mDurLegada

            info = "Dura" & ChrW(231) & ChrW(227) & "o mantida da opera" & ChrW(231) & ChrW(227) & "o: " & _
                   APS_FmtDur(dur) & vbCr & _
                   "(informe a quantidade para recalcular)"

            ObterDuracao = True

        Else

            info = "Informe a quantidade do lote em comprimidos."

        End If

    Else

        velBase = 0
        oee = 0

        d = APS_ParaDuracao(txtDur.Text)

        If d > 0 Then

            dur = d

            info = "Dura" & ChrW(231) & ChrW(227) & "o manual: " & APS_FmtDur(dur)

            If comp > 0 Then
                info = info & vbCr & "Quantidade: " & APS_FmtQtd(comp) & " comprimidos"
                cpc = APS_ParCPC(maq, cod)
                If cpc > 0 Then
                    info = info & " (" & APS_FmtQtd(Round(comp / cpc, 2)) & " caixas)"
                End If
            End If

            If APS_EhFette(maq) And Not mLegadoManual Then
                info = info & vbCr & "(FETTE sem velocidade cadastrada: cadastre em Cadastros e Par" & ChrW(226) & "metros para calcular a dura" & ChrW(231) & ChrW(227) & "o.)"
            End If

            ObterDuracao = True

        Else

            info = "Informe a dura" & ChrW(231) & ChrW(227) & "o (ex.: 10:30 ou 8h)."

        End If

    End If
End Function

' Minutos da preparacao escolhida (0 = Nenhuma / sem parametro valido)
Private Function SetupSelecionadoMin() As Long
    Dim rot() As String, minu() As Long, kk As Long, ii As Long, r As String

    SetupSelecionadoMin = 0
    If optSetupSim Is Nothing Or cboSetupParam Is Nothing Then Exit Function
    If Not optSetupSim.Value Then Exit Function
    r = Trim$(cboSetupParam.Text)
    If Len(r) = 0 Then Exit Function
    kk = APS_ParametrosSetup(cboMaquina.Text, rot, minu)
    For ii = 1 To kk
        If APS_Ig(rot(ii), r) Then SetupSelecionadoMin = minu(ii): Exit Function
    Next ii
End Function

' Inicio da PRODUCAO quando a preparacao comeca em inicioPrep: so conta tempo de calendario (Setup + Limpeza).
' Ex.: preparacao 10:00 + setup 1:30 = producao 11:30. Sem preparacao, devolve o proprio horario.
Private Function IniProducao(ByVal inicioPrep As Double, ByVal setupMin As Long) As Double
    Dim t As Double

    IniProducao = inicioPrep
    If setupMin <= 0 Then Exit Function
    t = APS_ProximoDisponivel(inicioPrep)
    If t < 0 Then Exit Function
    t = APS_AdicionarTrabalho(t, setupMin / 1440#)
    If t < 0 Then Exit Function
    t = APS_ProximoDisponivel(t)
    If t < 0 Then Exit Function
    IniProducao = APS_ArredMin(t)
End Function

' Resumo do cadastro de parametros (lote padrao, comprimidos por caixa, velocidade) lido SEMPRE do cadastro salvo.
' Sem cadastro: informa claramente (em vermelho) e a duracao/quantidade seguem manuais; nada e calculado com valor assumido.
Private Sub AtualizarCadastro()
    Dim s As String, falta As Boolean

    If lblCad Is Nothing Or cboCodigo Is Nothing Or cboMaquina Is Nothing Then Exit Sub

    If Len(Trim$(cboCodigo.Text)) = 0 Or Len(cboMaquina.Text) = 0 Then
        lblCad.Caption = ""
        Exit Sub
    End If

    s = APS_ResumoCadastro(cboMaquina.Text, Trim$(cboCodigo.Text), falta)

    If falta Then
        lblCad.ForeColor = RGB(192, 0, 0)
    Else
        lblCad.ForeColor = RGB(90, 90, 90)
    End If

    lblCad.Caption = s
End Sub

' Protege contra chamadas repetidas/aninhadas (eventos dos controles) e nunca deixa uma falha da previa travar o formulario
Private Sub Recalcular()
    If mRecalc Then Exit Sub
    mRecalc = True
    On Error GoTo Falha
    RecalcularInterno
    mRecalc = False
    Exit Sub
Falha:
    mRecalc = False
    On Error Resume Next
    lblAviso.ForeColor = RGB(192, 0, 0)
    lblAviso.Caption = "Falha ao calcular a pr" & ChrW(233) & "via: " & Err.Description
End Sub

Private Sub RecalcularInterno()
    Dim dur As Double
    Dim info As String
    Dim velBase As Double
    Dim oee As Double
    Dim d As Double
    Dim h As Double
    Dim aviso As String
    Dim t As String

    Dim tmp() As tOperacao
    Dim mudou() As Boolean
    Dim calOk As Boolean

    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim qtd As Long
    Dim lista As String
    Dim stMin As Long, pIni As Double, pa As Long

    ' a tela ainda esta sendo montada (eventos de Setup disparam antes dos labels existirem)
    If lblCalc Is Nothing Or lblAviso Is Nothing Or cboMaquina Is Nothing Then Exit Sub
    If txtData Is Nothing Or txtInicio Is Nothing Or txtAtraso Is Nothing Then Exit Sub

    AtualizarCadastro

    If Len(cboMaquina.Text) = 0 Then
        lblCalc.Caption = ""
        lblAviso.Caption = ""
        Exit Sub
    End If

    d = APS_ParaData(txtData.Text)
    h = APS_ParaHora(txtInicio.Text)

    If ObterDuracao(dur, info, velBase, oee) Then

        t = info

        If DurAuto() Then
            mBloq = mBloq + 1

            txtDur.Text = Format$(CLng(Round(dur * 1440#, 0)) \ 60, "0") & ":" & _
                          Format$(CLng(Round(dur * 1440#, 0)) Mod 60, "00")

            mBloq = mBloq - 1
        End If

    Else

        t = info

    End If

    If dur > 0 And d >= 0 And h >= 0 Then

        n = mN

        ReDim tmp(1 To n + 2)

        For i = 1 To n

            tmp(i) = mOps(i)

            If APS_Ig(tmp(i).id, mID) Then
                k = i
            End If

        Next i

        If k = 0 Then
            n = n + 1
            k = n
        End If

        ' previa igual ao salvar: nova OP ou inicio/duracao/maquina alterados = posicao escolhida
        Dim prioEsc As Boolean, atrP As Long
        If Not ParaAtraso(txtAtraso.Text, atrP) Then atrP = 0
        stMin = SetupSelecionadoMin()
        pIni = IniProducao(APS_ArredMin(d + h), stMin)       ' inicio da PRODUCAO (= inicio da preparacao + setup)
        If mNovo Or APS_Ig(tmp(k).id, "") Then
            prioEsc = True
        Else
            prioEsc = (Abs(pIni - tmp(k).ini) > APS_EPS) Or (Abs(dur + atrP / 1440# - tmp(k).dur) > APS_EPS) Or _
                      (Not APS_Ig(cboMaquina.Text, tmp(k).maquina))
        End If

        tmp(k).id = mID
        tmp(k).codigo = Trim$(cboCodigo.Text)
        tmp(k).Produto = txtProduto.Text
        tmp(k).lote = Trim$(cboLote.Text)
        tmp(k).maquina = cboMaquina.Text
        tmp(k).ini = pIni
        tmp(k).dur = APS_ArredMin(dur + atrP / 1440#)        ' total programado = original + atraso
        tmp(k).atraso = APS_ArredMin(atrP / 1440#)
        tmp(k).Fim = APS_ArredMin(tmp(k).ini + tmp(k).dur)
        tmp(k).status = cboStatus.Text
        tmp(k).SemDur = False

        ' Preparacao em memoria (igual ao salvar): atualiza/cria o bloco Setup ou cancela se "Nenhuma"
        pa = APS_AcharAuxDaProducao(tmp, n, mID, APS_TIPO_SETUP)
        If pa = 0 Then pa = APS_AcharAuxDaProducao(tmp, n, mID, APS_TIPO_LIMPEZA)
        If stMin > 0 Then
            If pa > 0 Then
                tmp(pa).dur = stMin / 1440#
                tmp(pa).SemDur = False
                tmp(pa).Fim = tmp(k).ini
                tmp(pa).ini = APS_SubtrairTrabalho(tmp(pa).Fim, tmp(pa).dur)
            Else
                n = n + 1
                APS_MontarAuxDaColuna tmp(n), tmp(k), stMin / 1440#
            End If
        ElseIf pa > 0 Then
            tmp(pa).status = APS_ST_CANCELADA
        End If

        APS_CalRecarregar

        qtd = APS_Encaixar(tmp, n, k, mudou, calOk, False, False, prioEsc)

        If calOk Then

            pa = APS_AcharAuxDaProducao(tmp, n, mID, APS_TIPO_SETUP)
            If pa = 0 Then pa = APS_AcharAuxDaProducao(tmp, n, mID, APS_TIPO_LIMPEZA)
            If stMin > 0 And pa > 0 Then
                t = t & vbCr & vbCr & "Prepara" & ChrW(231) & ChrW(227) & "o: " & Format$(tmp(pa).ini, "dd/mm/yyyy hh:nn") & " a " & Format$(tmp(pa).Fim, "hh:nn")
            Else
                t = t & vbCr
            End If
            t = t & vbCr & _
                "In" & ChrW(237) & "cio da produ" & ChrW(231) & ChrW(227) & "o: " & Format$(tmp(k).ini, "dd/mm/yyyy hh:nn") & vbCr & _
                "Fim: " & Format$(tmp(k).Fim, "dd/mm/yyyy hh:nn")

            If Abs(tmp(k).ini - pIni) > APS_EPS Then
                aviso = "In" & ChrW(237) & "cio ajustado pelo calend" & ChrW(225) & "rio / m" & _
                        ChrW(225) & "quina ocupada (pedido: " & _
                        Format$(pIni, "dd/mm hh:nn") & ")." & vbCr
            End If

            If qtd > 0 Then
                aviso = aviso & qtd & " opera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & _
                        ") seguinte(s) ser" & ChrW(227) & "o deslocada(s) em cascata." & vbCr
            End If

            APS_DetectarConflitos tmp, n, mudou

            lista = APS_ListaConflitos(tmp, n, k)

            If Len(lista) > 0 Then

                aviso = aviso & "CONFLITO com opera" & ChrW(231) & ChrW(227) & _
                        "o que n" & ChrW(227) & "o pode ser movida: " & lista

                lblAviso.ForeColor = RGB(192, 0, 0)

            Else

                If Len(aviso) = 0 Then
                    aviso = "Sem conflito na m" & ChrW(225) & "quina."
                End If

                lblAviso.ForeColor = RGB(0, 110, 40)

            End If

        Else

            If Len(APS_MotivoEncaixe) > 0 Then
                aviso = APS_MotivoEncaixe
            Else
                aviso = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o tem per" & _
                        ChrW(237) & "odo de trabalho."
            End If

            lblAviso.ForeColor = RGB(192, 0, 0)

        End If

    Else

        lblAviso.ForeColor = RGB(192, 0, 0)

    End If

    lblCalc.Caption = t
    lblAviso.Caption = aviso
End Sub

'----------------------------------------------------------
' Salvar
'----------------------------------------------------------
Private Function Validar(ByRef o As tOperacao, ByRef msg As String, ByRef temSetup As Boolean, _
                          ByRef setupRotulo As String, ByRef setupMin As Long) As Boolean
    Dim cod As String
    Dim lote As String
    Dim outro As String
    Dim info As String
    Dim rot() As String, minu() As Long, kk As Long, ii As Long

    Dim d As Double
    Dim h As Double
    Dim dur As Double
    Dim velBase As Double
    Dim oee As Double

    Dim p1 As Double
    Dim p2 As Double
    Dim atrMin As Long
    Dim cpcx As Double

    cod = Trim$(cboCodigo.Text)

    If Len(txtProduto.Text) = 0 Then
        msg = "Informe um c" & ChrW(243) & "digo de produto cadastrado."
        Exit Function
    End If

    lote = Trim$(cboLote.Text)

    If Len(lote) = 0 Then
        msg = "Informe o lote."
        Exit Function
    End If

    outro = APS_CodigoDoLote(mOps, mN, lote, mID)

    If Len(outro) > 0 And outro <> cod Then
        msg = "O lote " & lote & " j" & ChrW(225) & " est" & ChrW(225) & _
              " cadastrado para o produto " & outro & " " & ChrW(8212) & _
              " " & APS_NomeDoCodigo(outro) & "."
        Exit Function
    End If

    If Len(cboMaquina.Text) = 0 Then
        msg = "Selecione a m" & ChrW(225) & "quina."
        Exit Function
    End If

    If Not APS_ProdutoPermitido(cod, cboMaquina.Text) Then
        msg = "Este produto n" & ChrW(227) & "o pode ser programado nesta m" & ChrW(225) & "quina."
        Exit Function
    End If

    d = APS_ParaData(txtData.Text)

    If d < 0 Then
        msg = "Data inv" & ChrW(225) & "lida (use dd/mm/aaaa)."
        Exit Function
    End If

    h = APS_ParaHora(txtInicio.Text)

    If h < 0 Then
        msg = "Hor" & ChrW(225) & "rio inv" & ChrW(225) & "lido (use hh:mm)."
        Exit Function
    End If

    If Not ObterDuracao(dur, info, velBase, oee) Then
        msg = Replace(info, vbCr, " ")
        Exit Function
    End If

    If Not ParaAtraso(txtAtraso.Text, atrMin) Then
        msg = "Atraso inv" & ChrW(225) & "lido. Use hh:mm (ex.: 2:00), at" & ChrW(233) & " 240:00, ou deixe vazio."
        Exit Function
    End If

    If APS_DiasVisiveis(p1, p2) Then

        If d < p1 Or d > p2 Then

            If MsgBox("A data " & Format$(d, "dd/mm/yyyy") & _
                      " est" & ChrW(225) & " fora do per" & ChrW(237) & _
                      "odo exibido (" & Format$(p1, "dd/mm") & _
                      " a " & Format$(p2, "dd/mm") & ")." & vbCrLf & _
                      "Use + ADICIONAR DIAS para exibi-la. Salvar mesmo assim?", _
                      vbYesNo + vbQuestion, APS_TITULO) = vbNo Then

                msg = ""
                Exit Function

            End If

        End If

    End If

    o.linha = mLinha
    o.id = mID
    o.codigo = cod
    o.Produto = txtProduto.Text
    o.lote = lote
    o.ordem = Trim$(txtOP.Text)   ' OP: campo proprio, independente do ID, do lote e do codigo do produto
    o.maquina = cboMaquina.Text
    o.ini = APS_ArredMin(d + h)
    o.atraso = APS_ArredMin(atrMin / 1440#)
    o.dur = APS_ArredMin(dur + o.atraso)          ' total programado; a coluna Duracao grava so a ORIGINAL (APS_DurOriginal)
    o.Fim = APS_ArredMin(o.ini + o.dur)
    ' quantidade OFICIAL em comprimidos; caixas = quantidade derivada (so para os calculos e demais modulos)
    o.comp = APS_ParaQtd(txtQtd.Text)
    If o.comp > 0 Then
        cpcx = APS_ParCPC(cboMaquina.Text, cod)
        If cpcx <= 0 Then cpcx = APS_CP_CAIXA      ' sem conteudo por caixa cadastrado: constante ja existente (so informativo)
        o.caixas = o.comp / cpcx
    Else
        o.comp = 0
        o.caixas = 0
    End If
    o.status = cboStatus.Text

    If Len(o.status) = 0 Then
        o.status = APS_ST_PLANEJADA
    End If

    o.obs = Trim$(txtObs.Text)
    o.oee = oee
    o.velBase = velBase
    o.SemDur = False

    temSetup = optSetupSim.Value
    setupRotulo = Trim$(cboSetupParam.Text)
    setupMin = 0

    If temSetup Then
        If Len(setupRotulo) = 0 Then
            msg = "Selecione o par" & ChrW(226) & "metro da prepara" & ChrW(231) & ChrW(227) & "o (ou marque Nenhuma)."
            Exit Function
        End If
        kk = APS_ParametrosSetup(cboMaquina.Text, rot, minu)
        For ii = 1 To kk
            If APS_Ig(rot(ii), setupRotulo) Then setupMin = minu(ii): Exit For
        Next ii
        If setupMin <= 0 Then
            msg = "Par" & ChrW(226) & "metro de prepara" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lido para esta m" & ChrW(225) & "quina."
            Exit Function
        End If
        ' o horario digitado e o INICIO DA PREPARACAO: a producao comeca quando o Setup + Limpeza terminam
        o.ini = IniProducao(APS_ArredMin(d + h), setupMin)
        o.Fim = APS_ArredMin(o.ini + o.dur)
    End If

    Validar = True
End Function

' Atraso digitado -> minutos. Vazio, 0 ou 0:00 = sem atraso (0). Retorna False se o texto for invalido.
Private Function ParaAtraso(ByVal s As String, ByRef minutos As Long) As Boolean
    Dim d As Double

    minutos = 0
    s = Trim$(s)
    If Len(s) = 0 Or s = "0" Or s = "0:00" Or s = "00:00" Then ParaAtraso = True: Exit Function
    d = APS_ParaDuracao(s)
    If d <= 0 Then Exit Function
    minutos = CLng(Round(d * 1440#, 0))
    If minutos > 14400 Then minutos = 0: Exit Function      ' ate 10 dias
    ParaAtraso = True
End Function

Private Sub cmdSalvar_Click()
    Dim o As tOperacao
    Dim msg As String
    Dim temSetup As Boolean, setupRotulo As String, setupMin As Long

    If Not Validar(o, msg, temSetup, setupRotulo, setupMin) Then

        If Len(msg) > 0 Then
            MsgBox msg, vbExclamation, APS_TITULO
        End If

        Exit Sub
    End If

    If mTinhaSetupAoAbrir And Not temSetup Then
        If MsgBox("Esta opera" & ChrW(231) & ChrW(227) & "o possui uma prepara" & ChrW(231) & ChrW(227) & "o (Setup + Limpeza) vinculada. Deseja remov" & _
                  ChrW(234) & "-la?", vbYesNo + vbQuestion + vbDefaultButton2, APS_TITULO) <> vbYes Then
            Exit Sub
        End If
    End If

    If temSetup Or mTinhaSetupAoAbrir Then

        If APS_SalvarOperacaoComSetup(o, mNovo, temSetup, setupRotulo, setupMin, msg) Then
            If Len(msg) > 0 Then MsgBox msg, vbInformation, APS_TITULO
            Unload Me
        Else
            If Len(msg) > 0 Then MsgBox msg, vbExclamation, APS_TITULO
        End If

    Else

        If APS_SalvarOperacao(o, mNovo) Then
            Unload Me
        End If

    End If

End Sub

'----------------------------------------------------------
' EXCLUIR OPERACAO
'----------------------------------------------------------
Private Sub cmdExcluir_Click()
    Dim resposta As VbMsgBoxResult

    If mNovo Then Exit Sub

    If Len(Trim$(mID)) = 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o inv" & ChrW(225) & "lida para exclus" & ChrW(227) & "o.", _
               vbExclamation, APS_TITULO
        Exit Sub
    End If

    resposta = MsgBox( _
        "Tem certeza que deseja excluir esta opera" & ChrW(231) & ChrW(227) & "o?" & vbCrLf & vbCrLf & _
        "OP: " & IIf(Len(Trim$(txtOP.Text)) > 0, Trim$(txtOP.Text), "(sem OP)") & vbCrLf & _
        "Produto: " & txtProduto.Text & vbCrLf & _
        "Lote: " & cboLote.Text & vbCrLf & _
        "M" & ChrW(225) & "quina: " & cboMaquina.Text & vbCrLf & vbCrLf & _
        "Esta opera" & ChrW(231) & ChrW(227) & "o ser" & ChrW(225) & " removida do planejamento.", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        APS_TITULO & " - Excluir")

    If resposta <> vbYes Then Exit Sub

    If APS_ExcluirOperacao(mID) Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
