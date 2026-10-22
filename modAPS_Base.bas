Attribute VB_Name = "modAPS_Base"
Option Explicit

'==========================================================
' APS PURAN - BASE
' Constantes, tipos, cadastro (produtos/maquinas), configuracao
' gravada no proprio arquivo, leitura/gravacao de 02_Operacoes
' e migracao do layout antigo.
' O sistema usa somente 02_Operacoes e 1_Planejamento.
'
' ESTRUTURA ATUAL DE 02_Operacoes:
' 1  Codigo do Produto
' 2  Produto
' 3  Lote
' 4  Maquina
' 5  Data de Inicio
' 6  Data de Fim
' 7  Planejamento
' 8  Observacoes
' 9  Quantidade de Caixas
' 10 Duracao
'
' ID, OEE e Velocidade-base NAO sao colunas da planilha.
'==========================================================

Public Const APS_ABA_OPS As String = "02_Operacoes"
Public Const APS_ABA_PLAN As String = "1_Planejamento"

Public Const APS_TITULO As String = "APS Puran"
Public Const APS_EPS As Double = 0.0000005

Public Const APS_CP_CAIXA As Double = 30

' Status
Public Const APS_ST_PLANEJADA As String = "Planejada"
Public Const APS_ST_ANDAMENTO As String = "Em Andamento"
Public Const APS_ST_ATRASADA As String = "Atrasada"
Public Const APS_ST_CANCELADA As String = "Cancelada"

Public Const APS_ST_CONFLITO As String = "Conflito"

' Parametros
Public Const APS_MED_SETUP_A_MIN As Long = 15
Public Const APS_MED_SETUP_B_MIN As Long = 90
Public Const APS_MED_SETUP_D_MIN As Long = 120
Public Const APS_MED_SETUP_E_MIN As Long = 300
Public Const APS_MED_CAMPANHA_LOTES As Long = 12
Public Const APS_FET_LIMPEZA_TOTAL_MIN As Long = 720
Public Const APS_FET_LIMPEZA_PARCIAL_MIN As Long = 600
Public Const APS_FET_TROCA_LOTE_MIN As Long = 30
Public Const APS_FET_CAMPANHA_LOTES As Long = 8
Public Const APS_SALA_LIMPEZA_PARCIAL_MIN As Long = 30
Public Const APS_SALA_LIMPEZA_TOTAL_MIN As Long = 480
Public Const APS_SALA_CAMPANHA_LOTES As Long = 11

' Setup / Limpeza automaticos (linhas de 02_Operacoes identificadas pelo codigo)
Public Const APS_COD_SETUP As String = "SETUP"
Public Const APS_COD_LIMPEZA As String = "LIMPEZA"

' Tipos de linha vinculada. Os codigos APS_COD_* acima sao SOMENTE o formato ANTIGO
' (registros ja gravados com Codigo = SETUP): continuam reconhecidos, nao sao mais gravados.
Public Const APS_TIPO_SETUP As String = "SETUP"
Public Const APS_TIPO_LIMPEZA As String = "LIMPEZA"

' Cor do cabecalho de 02_Operacoes (RGB 31,78,120) - usada so para detectar/limpar
' formatacao de cabecalho que foi copiada para linhas de dados.
Public Const APS_COR_CAB As Long = 7884319

Public Type tOperacao
    linha As Long
    id As String
    codigo As String
    Produto As String
    lote As String
    ordem As String      ' OP (Ordem de Producao): campo PROPRIO da operacao. NAO e o ID tecnico, nem a linha, nem o lote.
    maquina As String
    ini As Double
    Fim As Double
    dur As Double
    caixas As Double
    comp As Double       ' QUANTIDADE OFICIAL em COMPRIMIDOS (0 = operacao antiga, so com caixas). caixas = comp / comprimidos por caixa
    status As String
    obs As String
    oee As Double
    velBase As Double
    SemDur As Boolean
    tipoAux As String    ' "" = producao; "SETUP"/"LIMPEZA" = linha vinculada a uma producao (NAO e uma OP)
    idPai As String      ' ID da producao a que a linha SETUP/LIMPEZA pertence
    rotAux As String     ' rotulo do parametro escolhido (ex.: "Setup D")
    atraso As Double     ' tempo ADICIONAL de programacao (dias), no fim da producao. dur (em memoria) = duracao original + atraso
End Type

'----------------------------------------------------------
' [CORRECAO] Tipo de backup de linha. Declarado ANTES do primeiro
' procedimento (junto a tOperacao) para que o VBA consiga compilá-lo.
' O erro "O tipo definido pelo usuário não foi definido" ocorria porque
' este Type estava declarado na linha 3338, DEPOIS do primeiro procedimento
' do modulo (APS_ST_CONCLUIDA, linha 113). Em VBA, todos os Type devem
' constar antes do primeiro Sub/Function do modulo.
'----------------------------------------------------------
Private Type tBackupLinha
    linha As Long
    nCol As Long
    valores() As Variant
    formatos() As String
End Type

Public APS_OpsIgnoradas As Long

' Trava contra reentrada dos eventos da planilha (>0 = o proprio APS esta gravando)
Public APS_Ocupado As Long

' Ultimo retrato conhecido de 02_Operacoes e dos cards (usado por APS_PrecisaAtualizar / APS_Tick). Tipos do projeto antigo.
Public APS_UltAssin As Double
Public APS_UltCards As Long

' Protecao mensal: a cascata nunca sai do mes da operacao que esta sendo encaixada.
' APS_ForaDoMes = True quando nao coube; APS_MesRef = ultimo dia do mes limite. (Declaracao do projeto antigo.)
' Declaracoes de modulo precisam ficar ANTES do primeiro procedimento.
Public APS_ForaDoMes As Boolean
' Motivo (texto) quando APS_Encaixar recusa a posicao escolhida por uma restricao real (modo "escolha do usuario").
Public APS_MotivoEncaixe As String
Public APS_MesRef As Double

'----------------------------------------------------------
' Textos com acento
'----------------------------------------------------------
Public Function APS_ST_CONCLUIDA() As String
    APS_ST_CONCLUIDA = "Conclu" & ChrW(237) & "da"
End Function

Public Function H_COD() As String
    H_COD = "C" & ChrW(243) & "digo do Produto"
End Function

Public Function H_PRO() As String
    H_PRO = "Produto"
End Function

' Cabecalho da coluna OP (Ordem de Producao). Fica sempre logo depois de Produto.
Public Function H_ORD() As String
    H_ORD = "OP"
End Function

Public Function H_LOT() As String
    H_LOT = "Lote"
End Function

Public Function H_MAQ() As String
    H_MAQ = "M" & ChrW(225) & "quina"
End Function

Public Function H_INI() As String
    H_INI = "Data de In" & ChrW(237) & "cio"
End Function

Public Function H_FIM() As String
    H_FIM = "Data de Fim"
End Function

Public Function H_PLA() As String
    H_PLA = "Planejamento"
End Function

Public Function H_OBS() As String
    H_OBS = "Observa" & ChrW(231) & ChrW(245) & "es"
End Function

Public Function H_QTD() As String
    H_QTD = "Quantidade de Caixas"
End Function

' Quantidade oficial do lote em COMPRIMIDOS. Coluna criada so quando a primeira operacao com quantidade em comprimidos
' e gravada; operacoes antigas ficam com a celula vazia (nada e reescrito) e continuam usando 'Quantidade de Caixas'.
Public Function H_QCP() As String
    H_QCP = "Quantidade (comprimidos)"
End Function

' Comprimidos da operacao: a quantidade oficial (coluna Quantidade (comprimidos)) ou, em operacao antiga so com
' caixas, caixas x APS_CP_CAIXA (a conversao que o Painel ja usava). Nao grava nada.
Public Function APS_CompDaOp(ByRef o As tOperacao) As Double
    If o.comp > 0 Then
        APS_CompDaOp = o.comp
    Else
        APS_CompDaOp = o.caixas * APS_CP_CAIXA
    End If
End Function

Public Function H_DUR() As String
    H_DUR = "Dura" & ChrW(231) & ChrW(227) & "o"
End Function

' Mantidas para compatibilidade interna com outros modulos.
' NAO sao mais usadas como colunas da 02_Operacoes.
Public Function H_ID() As String
    H_ID = "ID"
End Function

Public Function H_OEE() As String
    H_OEE = "OEE"
End Function

Public Function H_VEL() As String
    H_VEL = "Velocidade-base"
End Function

' Cabecalho da coluna Setup: tempo de Setup/Limpeza DA PROPRIA operacao. O Setup NAO tem linha propria em
' 02_Operacoes; o bloco do Setup (e o card) e montado a partir desta coluna em APS_LerOps.
Public Function H_SET() As String
    H_SET = "Setup"
End Function

' Cabecalho da coluna Atraso: tempo adicional de programacao da operacao (no fim da producao). A coluna Duracao
' continua sendo a duracao ORIGINAL; a coluna Fim e o termino programado (duracao + atraso, pelo calendario).
Public Function H_ATR() As String
    H_ATR = "Atraso"
End Function

' Duracao ORIGINAL de producao (a que vai para a coluna Duracao): total em memoria menos o atraso.
Public Function APS_DurOriginal(ByRef o As tOperacao) As Double
    If o.atraso > APS_EPS And o.dur - o.atraso > APS_EPS Then
        APS_DurOriginal = APS_ArredMin(o.dur - o.atraso)
    Else
        APS_DurOriginal = o.dur
    End If
End Function

Public Function APS_ListaStatus() As Variant
    APS_ListaStatus = Array(APS_ST_PLANEJADA, APS_ST_ANDAMENTO, APS_ST_CONCLUIDA(), _
                            APS_ST_ATRASADA, APS_ST_CANCELADA, APS_ST_CONFLITO)
End Function

Public Function APS_Ig(ByVal a As String, ByVal b As String) As Boolean
    APS_Ig = (StrComp(Trim$(a), Trim$(b), vbTextCompare) = 0)
End Function

' Operacao ativa = nao Cancelada (mesma regra do projeto antigo).
Public Function APS_Ativa(ByRef o As tOperacao) As Boolean
    APS_Ativa = Not APS_Ig(o.status, APS_ST_CANCELADA)
End Function

Public Function APS_Txt(ByVal v As Variant) As String
    If IsError(v) Then
        APS_Txt = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        APS_Txt = ""
    Else
        APS_Txt = Trim$(CStr(v))
    End If
End Function

'----------------------------------------------------------
' Abas e protecao
'----------------------------------------------------------
Public Function APS_Aba(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set APS_Aba = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
End Function

Public Function APS_Liberar(ByVal ws As Worksheet) As Boolean
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0

        If ws.ProtectContents Or ws.ProtectDrawingObjects Then
            Err.Raise vbObjectError + 513, , _
                "A aba '" & ws.Name & "' est" & ChrW(225) & " protegida com senha."
        End If

        APS_Liberar = True
    End If
End Function

Public Sub APS_Reproteger(ByVal ws As Worksheet, ByVal estava As Boolean)
    If estava Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Public Sub APS_ProtegerAbas()
    Dim ws As Worksheet

    Set ws = APS_Aba(APS_ABA_OPS)
    If Not ws Is Nothing Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

    Set ws = APS_Aba(APS_ABA_PLAN)
    If Not ws Is Nothing Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

    MsgBox "Abas 02_Operacoes e 1_Planejamento protegidas (sem senha).", _
           vbInformation, APS_TITULO
End Sub

Public Sub APS_DesprotegerAbas()
    Dim ws As Worksheet

    Set ws = APS_Aba(APS_ABA_OPS)
    If Not ws Is Nothing Then ws.Unprotect

    Set ws = APS_Aba(APS_ABA_PLAN)
    If Not ws Is Nothing Then ws.Unprotect

    MsgBox "Abas desprotegidas.", vbInformation, APS_TITULO
End Sub

Public Function APS_Col(ByVal ws As Worksheet, ByVal cab As String) As Long
    Dim c As Long, ult As Long

    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ult = 1 And Len(APS_Txt(ws.Cells(1, 1).Value)) = 0 Then Exit Function

    For c = 1 To ult
        If StrComp(APS_Txt(ws.Cells(1, c).Value), cab, vbTextCompare) = 0 Then
            APS_Col = c
            Exit Function
        End If
    Next c
End Function

'----------------------------------------------------------
' Configuracao
'----------------------------------------------------------
Public Function APS_CfgLer(ByVal chave As String, ByVal padrao As String) As String
    Dim i As Long, t As String, s As String, achou As Boolean

    i = 1

    Do
        t = ""

        On Error Resume Next
        t = ThisWorkbook.Names("APS_CFG_" & chave & "_" & i).RefersTo
        On Error GoTo 0

        If Len(t) < 3 Then Exit Do

        achou = True
        s = s & mID$(t, 3, Len(t) - 3)
        i = i + 1
    Loop

    If achou Then
        APS_CfgLer = s
    Else
        APS_CfgLer = padrao
    End If
End Function

Public Sub APS_CfgGravar(ByVal chave As String, ByVal valor As String)
    Dim i As Long, n As Long, p As Long, nm As String

    valor = Replace(valor, """", "")
    valor = Replace(valor, vbCrLf, "~")
    valor = Replace(valor, vbCr, "~")
    valor = Replace(valor, vbLf, "~")

    i = 1

    Do
        nm = "APS_CFG_" & chave & "_" & i

        On Error Resume Next
        Err.Clear
        ThisWorkbook.Names(nm).Delete

        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If

        On Error GoTo 0
        i = i + 1
    Loop

    n = 1
    p = 1

    Do
        ThisWorkbook.Names.Add _
            Name:="APS_CFG_" & chave & "_" & n, _
            RefersTo:="=""" & mID$(valor, p, 200) & """", _
            Visible:=False

        p = p + 200
        n = n + 1
    Loop While p <= Len(valor)
End Sub

'----------------------------------------------------------
' Conversoes
'----------------------------------------------------------
Public Function APS_ParaData(ByVal v As Variant) As Double
    Dim s As String, p() As String
    Dim d As Long, m As Long, a As Long

    APS_ParaData = -1

    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        APS_ParaData = Int(CDbl(v) + APS_EPS)
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        If Len(s) = 0 Then Exit Function

        If InStr(s, " ") > 0 Then
            s = Left$(s, InStr(s, " ") - 1)
        End If

        s = Replace(s, "-", "/")
        s = Replace(s, ".", "/")

        p = Split(s, "/")

        If UBound(p) <> 2 Then Exit Function

        If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function

        d = CLng(p(0))
        m = CLng(p(1))
        a = CLng(p(2))

        If d > 31 Then
            a = d
            d = CLng(p(2))
        End If

        If a < 100 Then a = a + 2000

        If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
        If Day(DateSerial(a, m, d)) <> d Then Exit Function

        APS_ParaData = CDbl(DateSerial(a, m, d))
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) >= 1 Then
            APS_ParaData = Int(CDbl(v) + APS_EPS)
        End If
    End If
End Function

Public Function APS_ParaHora(ByVal v As Variant) As Double
    Dim s As String, p() As String
    Dim h As Long, m As Long, x As Double

    APS_ParaHora = -1

    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        x = CDbl(v)
        APS_ParaHora = x - Int(x)
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = LCase$(Trim$(CStr(v)))

        If Len(s) = 0 Then Exit Function

        s = Replace(s, "h", ":")

        If Right$(s, 1) = ":" Then s = s & "00"

        If InStr(s, ":") > 0 Then
            p = Split(s, ":")

            If Not IsNumeric(p(0)) Then Exit Function
            h = CLng(p(0))

            If UBound(p) >= 1 Then
                If Not IsNumeric(p(1)) Then Exit Function
                m = CLng(p(1))
            End If
        Else
            If s Like "*[!0-9]*" Then Exit Function

            Select Case Len(s)
                Case 1, 2
                    h = CLng(s)

                Case 3, 4
                    h = CLng(Left$(s, Len(s) - 2))
                    m = CLng(Right$(s, 2))

                Case Else
                    Exit Function
            End Select
        End If

        If h < 0 Or h > 23 Or m < 0 Or m > 59 Then Exit Function

        APS_ParaHora = (h * 60 + m) / 1440#
        Exit Function
    End If

    If IsNumeric(v) Then
        x = CDbl(v)

        If x >= 0 And x < 1 Then
            APS_ParaHora = x
        End If
    End If
End Function

Public Function APS_ParaDataHora(ByVal v As Variant) As Double
    Dim s As String, d As Double, h As Double, p As Long

    APS_ParaDataHora = -1

    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function

    If VarType(v) = vbDate Then
        APS_ParaDataHora = APS_ArredMin(CDbl(v))
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))

        d = APS_ParaData(s)
        If d < 0 Then Exit Function

        p = InStr(s, " ")

        If p > 0 Then
            h = APS_ParaHora(mID$(s, p + 1))
            If h < 0 Then Exit Function
        End If

        APS_ParaDataHora = d + h
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) >= 1 Then
            APS_ParaDataHora = APS_ArredMin(CDbl(v))
        End If
    End If
End Function

Public Function APS_ParaDuracao(ByVal s As String) As Double
    Dim p() As String, h As Double, m As Double

    APS_ParaDuracao = -1

    s = LCase$(Trim$(s))

    If Len(s) = 0 Then Exit Function

    s = Replace(s, "h", ":")

    If Right$(s, 1) = ":" Then s = s & "00"

    If InStr(s, ":") > 0 Then
        p = Split(s, ":")

        If Not IsNumeric(p(0)) Then Exit Function
        h = Val(p(0))

        If UBound(p) >= 1 Then
            If Not IsNumeric(p(1)) Then Exit Function
            m = Val(p(1))
        End If

        If h < 0 Or m < 0 Or m > 59 Then Exit Function
        If h = 0 And m = 0 Then Exit Function

        APS_ParaDuracao = (h * 60 + m) / 1440#
    Else
        s = Replace(s, ",", ".")

        If s Like "*[!0-9.]*" Then Exit Function

        h = Val(s)

        If h <= 0 Then Exit Function

        APS_ParaDuracao = APS_ArredMin(h / 24#)
    End If
End Function

Public Function APS_ParaQtd(ByVal s As String) As Double
    s = Replace(Replace(Trim$(s), ".", ""), " ", "")

    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9,]*" Then Exit Function
    If InStr(s, ",") <> InStrRev(s, ",") Then Exit Function

    s = Replace(s, ",", ".")

    APS_ParaQtd = Val(s)
End Function

Public Function APS_FmtQtd(ByVal q As Double) As String
    If q = Int(q) Then
        APS_FmtQtd = Format$(q, "#,##0")
    Else
        APS_FmtQtd = Format$(q, "#,##0.000")
    End If
End Function

Public Function APS_ArredMin(ByVal x As Double) As Double
    APS_ArredMin = Round(x * 1440#, 0) / 1440#
End Function

Public Function APS_FmtDur(ByVal dias As Double) As String
    Dim mins As Long

    mins = CLng(Round(dias * 1440#, 0))

    APS_FmtDur = CStr(mins \ 60) & "h" & Format$(mins Mod 60, "00") & "min"
End Function

'----------------------------------------------------------
' Cadastro de produtos
'----------------------------------------------------------
Private Function TabProdutos() As Variant
    TabProdutos = Array( _
        "906664|Puran T4 12,5 mcg|833.333|150", _
        "906655|Puran T4 25 mcg|4000|200", _
        "906665|Puran T4 37,5 mcg|833.333|150", _
        "906654|Puran T4 50 mcg|4000|200", _
        "906666|Puran T4 62,5 mcg|833.333|150", _
        "906656|Puran T4 75 mcg|2666.667|200", _
        "906658|Puran T4 100 mcg|2000|200", _
        "906660|Puran T4 125 mcg|2000|200", _
        "906661|Puran T4 150 mcg|1333.333|200", _
        "906663|Puran T4 175 mcg|1333.333|200", _
        "906667|Puran T4 300 mcg|833.333|200", _
        "906390|Mistura para Puran|-1|0")
End Function

Public Function APS_Produtos(ByRef cods() As String, ByRef nomes() As String, _
                             ByRef lotes() As Double, ByRef vels() As Double) As Long
    Dim t As Variant, i As Long, p() As String, n As Long

    t = TabProdutos()
    n = UBound(t) + 1

    ReDim cods(1 To n)
    ReDim nomes(1 To n)
    ReDim lotes(1 To n)
    ReDim vels(1 To n)

    For i = 0 To UBound(t)
        p = Split(CStr(t(i)), "|")

        cods(i + 1) = p(0)
        nomes(i + 1) = p(1)
        lotes(i + 1) = Val(p(2))
        vels(i + 1) = Val(p(3))
    Next i

    APS_Produtos = n
End Function

Public Function APS_IndiceProduto(ByVal codigo As String) As Long
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_IndiceProduto = i
            Exit Function
        End If
    Next i
End Function

Public Function APS_NormProd(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " COMPRIMIDOS", "")
    s = Replace(s, ".", ",")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    APS_NormProd = Trim$(s)
End Function

Public Function APS_CodigoDoNome(ByVal nome As String) As String
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If APS_NormProd(n(i)) = APS_NormProd(nome) Then
            APS_CodigoDoNome = c(i)
            Exit Function
        End If
    Next i
End Function

Public Function APS_NomeDoCodigo(ByVal codigo As String) As String
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_NomeDoCodigo = n(i)
            Exit Function
        End If
    Next i
End Function

Public Function APS_VelBaseDoCodigo(ByVal codigo As String) As Double
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim i As Long, k As Long

    k = APS_Produtos(c, n, l, v)

    For i = 1 To k
        If Trim$(c(i)) = Trim$(codigo) Then
            APS_VelBaseDoCodigo = v(i)
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------
' Maquinas
'----------------------------------------------------------
Public Function APS_Maquinas(ByRef m() As String) As Long
    ReDim m(1 To 6)

    m(1) = "Sala de Mistura"
    m(2) = "FETTE 2090"
    m(3) = "FETTE 2"
    m(4) = "Mediseal"
    m(5) = "Blisterflex"
    m(6) = "Granula" & ChrW(231) & ChrW(227) & "o"

    APS_Maquinas = 6
End Function

Public Function APS_EhMediseal(ByVal maq As String) As Boolean
    APS_EhMediseal = APS_Ig(maq, "Mediseal")
End Function

Public Function APS_EhFette(ByVal maq As String) As Boolean
    APS_EhFette = APS_Ig(maq, "FETTE 2090") Or APS_Ig(maq, "FETTE 2")
End Function

Public Function APS_EhSala(ByVal maq As String) As Boolean
    APS_EhSala = APS_Ig(maq, "Sala de Mistura")
End Function

' Maquinas que possuem parametros de Setup / Limpeza cadastrados
Public Function APS_TemSetup(ByVal maq As String) As Boolean
    APS_TemSetup = APS_EhMediseal(maq) Or APS_EhFette(maq) Or APS_EhSala(maq)
End Function

Public Function APS_EhAutoCod(ByVal cod As String) As Boolean
    APS_EhAutoCod = APS_Ig(cod, APS_COD_SETUP) Or APS_Ig(cod, APS_COD_LIMPEZA)
End Function

' Linha Setup/Limpeza (vinculada a uma OP, ou no formato antigo). Depende de tipoAux,
' preenchido por APS_ClassificarAux (APS_LerOps).
Public Function APS_EhAuto(ByRef o As tOperacao) As Boolean
    APS_EhAuto = (Len(o.tipoAux) > 0)
End Function

'----------------------------------------------------------
' SETUP escolhido pelo usuario no formulario. A linha de Setup NAO e uma OP: usa Codigo/
' Produto/Lote/Maquina da OP pai e fica ligada a ela pelo marcador tecnico
' [[APS_SETUP_DE:<ID da OP>]] em Observacoes (o ID e o da coluna K, persistente).
' Usa SOMENTE os parametros/tempos ja cadastrados acima; nenhuma regra de qual usar
' e decidida aqui - quem escolhe e o usuario, no formulario.
'----------------------------------------------------------

' Parametros de SETUP validos para a maquina informada (sem LIMPEZA: fora de escopo por ora).
' Devolve 0 se a maquina nao tiver nenhum parametro de SETUP cadastrado.
' PREPARACAO = Setup + Limpeza, uma UNICA preparacao vinculada a OP (nao duas operacoes,
' nao duas duracoes somadas). Para cada maquina, a lista abaixo junta os tempos de SETUP
' e de LIMPEZA ja cadastrados nas constantes acima (nenhum tempo novo) - o usuario escolhe
' UM desses valores no formulario para representar a preparacao daquela OP.
Public Function APS_ParametrosSetup(ByVal maquina As String, ByRef rotulos() As String, _
                                    ByRef minutos() As Long) As Long
    Dim n As Long

    ReDim rotulos(1 To 6)
    ReDim minutos(1 To 6)

    If APS_EhMediseal(maquina) Then
        n = 4
        rotulos(1) = "Setup A": minutos(1) = APS_MED_SETUP_A_MIN
        rotulos(2) = "Setup B": minutos(2) = APS_MED_SETUP_B_MIN
        rotulos(3) = "Setup D": minutos(3) = APS_MED_SETUP_D_MIN
        rotulos(4) = "Setup E": minutos(4) = APS_MED_SETUP_E_MIN
    ElseIf APS_EhFette(maquina) Then
        n = 3
        rotulos(1) = "Troca de lote": minutos(1) = APS_FET_TROCA_LOTE_MIN
        rotulos(2) = "Limpeza parcial": minutos(2) = APS_FET_LIMPEZA_PARCIAL_MIN
        rotulos(3) = "Limpeza total": minutos(3) = APS_FET_LIMPEZA_TOTAL_MIN
    ElseIf APS_EhSala(maquina) Then
        n = 2
        rotulos(1) = "Limpeza parcial": minutos(1) = APS_SALA_LIMPEZA_PARCIAL_MIN
        rotulos(2) = "Limpeza total": minutos(2) = APS_SALA_LIMPEZA_TOTAL_MIN
    End If
    ' Qualquer outra maquina: sem parametro cadastrado ainda.

    APS_ParametrosSetup = n
End Function

Public Function APS_MarcadorVinculo(ByVal tipo As String, ByVal idPai As String) As String
    APS_MarcadorVinculo = "[[APS_" & tipo & "_DE:" & idPai & "]]"
End Function

Public Function APS_MarcadorSetup(ByVal idProd As String) As String
    APS_MarcadorSetup = APS_MarcadorVinculo(APS_TIPO_SETUP, idProd)
End Function

' Le, das Observacoes, o ID da OP pai e o tipo (SETUP/LIMPEZA) do vinculo. "" se nao houver marcador.
Public Function APS_IDPaiDoVinculo(ByVal obs As String, ByRef tipo As String) As String
    Dim marca As String, p1 As Long, p2 As Long

    tipo = ""
    marca = "[[APS_" & APS_TIPO_SETUP & "_DE:"
    p1 = InStr(1, obs, marca, vbTextCompare)
    If p1 > 0 Then
        tipo = APS_TIPO_SETUP
    Else
        marca = "[[APS_" & APS_TIPO_LIMPEZA & "_DE:"
        p1 = InStr(1, obs, marca, vbTextCompare)
        If p1 > 0 Then tipo = APS_TIPO_LIMPEZA
    End If
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(marca)
    p2 = InStr(p1, obs, "]]")
    If p2 = 0 Then
        tipo = ""
        Exit Function
    End If
    APS_IDPaiDoVinculo = Trim$(mID$(obs, p1, p2 - p1))
End Function

' Compatibilidade: ID da OP pai de um marcador SETUP.
Public Function APS_IDDoSetup(ByVal obs As String) As String
    Dim t As String, s As String

    s = APS_IDPaiDoVinculo(obs, t)
    If APS_Ig(t, APS_TIPO_SETUP) Then APS_IDDoSetup = s
End Function

' Rotulo do parametro (ex.: "Setup D"): texto que segue o marcador, ate " - <maquina>".
Public Function APS_RotuloDoVinculo(ByVal obs As String) As String
    Dim p As Long, q As Long, s As String

    p = InStr(1, obs, "]]")
    If p = 0 Then Exit Function
    s = Trim$(mID$(obs, p + 2))
    q = InStrRev(s, " - ")
    If q > 1 Then s = Trim$(Left$(s, q - 1))
    APS_RotuloDoVinculo = s
End Function

' Classifica uma linha lida de 02_Operacoes: producao ("") ou Setup/Limpeza vinculada.
' Formato novo: marcador em Observacoes. Formato antigo: Codigo = SETUP / LIMPEZA.
Public Sub APS_ClassificarAux(ByRef o As tOperacao)
    Dim t As String

    o.tipoAux = ""
    o.idPai = ""
    o.rotAux = ""

    o.idPai = APS_IDPaiDoVinculo(o.obs, t)
    If Len(o.idPai) > 0 Then
        o.tipoAux = t
        o.rotAux = APS_RotuloDoVinculo(o.obs)
        If Len(o.rotAux) = 0 And APS_EhAutoCod(o.codigo) Then o.rotAux = o.Produto
    ElseIf APS_Ig(o.codigo, APS_COD_SETUP) Then
        o.tipoAux = APS_TIPO_SETUP
        o.rotAux = o.Produto
    ElseIf APS_Ig(o.codigo, APS_COD_LIMPEZA) Then
        o.tipoAux = APS_TIPO_LIMPEZA
        o.rotAux = o.Produto
    End If
End Sub

' Indice (dentro de ops) da linha vinculada de um TIPO a uma producao. 0 se nao houver.
Public Function APS_AcharAuxDaProducao(ByRef ops() As tOperacao, ByVal n As Long, _
                                       ByVal idProd As String, ByVal tipo As String) As Long
    Dim i As Long

    If Len(idProd) = 0 Then Exit Function
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 Then
            If APS_Ig(ops(i).tipoAux, tipo) And APS_Ig(ops(i).idPai, idProd) Then
                APS_AcharAuxDaProducao = i
                Exit Function
            End If
        End If
    Next i
End Function

' Compatibilidade: SETUP vinculado a uma producao.
Public Function APS_AcharSetupDaProducao(ByRef ops() As tOperacao, ByVal n As Long, _
                                         ByVal idProd As String) As Long
    APS_AcharSetupDaProducao = APS_AcharAuxDaProducao(ops, n, idProd, APS_TIPO_SETUP)
End Function

'----------------------------------------------------------
' APS_SincronizarSetups - RECONSTRUCAO COMPATIVEL - NAO CODIGO ORIGINAL.
'
' Evidencia decisiva sobre o ESCOPO desta rotina: o comentario ja existente,
' logo acima de APS_ParametrosSetup, e explicito - "nenhuma regra de qual
' [SETUP] usar e decidida aqui - quem escolhe e o usuario, no formulario".
' Ou seja, CRIAR um vinculo de Setup/Limpeza novo, ou decidir QUAL tipo usar,
' e confirmado como fora do escopo de qualquer rotina automatica - isso so
' acontece pela tela (APS_SalvarOperacaoComSetup, ainda nao reconstruida).
'
' O que resta, e que EXISTE evidencia solida para reconstruir, e a
' MANUTENCAO de um vinculo que ja existe: se a producao teve seu ini
' alterado (por edicao manual, SAP, ou pelo proprio Stage 3 rodando para
' OUTRA maquina) sem que o Setup/Limpeza ja vinculado a ela tenha sido
' atualizado junto, o par fica com o "fim" do Setup diferente do "ini" da
' producao - quebrando a regra "Setup.fim = Producao.ini" que o proprio
' APS_FecharVaos (Stage 3, NAO ALTERADO) ja impoe e preserva quando ele
' mesmo move o bloco. A formula abaixo (iniPar = APS_SubtrairTrabalho(fimPar,
' durPar)) e uma copia LITERAL da mesma formula ja existente dentro de
' APS_FecharVaos para o caso "t = producao, parceiro = Setup/Limpeza" - nao
' foi inventado nenhum criterio de calendario novo.
'
' Assinatura CONFIRMADA PELO CODIGO pelo unico ponto de chamada
' (modAPS_Cards.APS_Atualizar): "If APS_SincronizarSetups() Then" - sem
' parametros, retorno Boolean (True = alguma linha foi re-sincronizada,
' fazendo a chamadora re-ler ops() em seguida).
'
' [ADAPTACAO modelo novo] Alem de ini/fim, acompanha maquina/produto/lote da producao e nao
' remaneja Setup/Limpeza Em Andamento/Concluido nem de producao cancelada.
' STATUS: Setup/Limpeza nao tem status proprio; cancelamento/historico vem da producao (ver APS_LerOps).
' Producao cancelada: o Setup/Limpeza deixa de ocupar a maquina, sem apagar nem gravar nada.
' NAO cria nem remove nenhum vinculo (isso e form/APS_LimparSetupsOrfaos).
' NAO decide qual tipo de Setup usar. NAO toca idPai/tipoAux/rotAux/Produto/
' maquina/duracao de nada. So corrige ini/fim do Setup/Limpeza JA vinculado,
' quando ele esta fora de posicao em relacao a producao.
'
' DESCONHECIDO, informado explicitamente: se a rotina original fazia mais do
' que isso (por exemplo, tambem validar limite mensal do novo horario do
' Setup, ou tratar Limpeza-depois-da-producao de forma diferente de
' Setup-antes-da-producao) nao ha evidencia no codigo atual nem no cache
' para confirmar ou descartar - o escopo abaixo e deliberadamente o minimo
' comprovavel, nao uma reconstrucao completa de tudo que o nome sugere.
Public Function APS_SincronizarSetups() As Boolean
    Dim ops() As tOperacao, n As Long, i As Long, p As Long
    Dim durPar As Double, iniPar As Double, fimPar As Double
    Dim difHora As Boolean, difDados As Boolean
    Dim j As Long, lin As Long, achou As Boolean
    Dim backups() As tBackupLinha, nBk As Long, iBk As Long
    Dim falhaRecuperou As Boolean, msg As String

    On Error GoTo Falha

    n = APS_LerOps(ops)
    If n <= 0 Then Exit Function

    ' [PEN-B-21] Fase 1: identificar quais linhas serao sincronizadas e fazer backup antes de qualquer gravacao.
    ReDim backups(1 To IIf(n < 1, 1, n))
    nBk = 0
    For i = 1 To n
        If Len(ops(i).tipoAux) = 0 Then   ' i e producao
            p = APS_AcharAuxDaProducao(ops, n, ops(i).id, APS_TIPO_SETUP)
            If p = 0 Then p = APS_AcharAuxDaProducao(ops, n, ops(i).id, APS_TIPO_LIMPEZA)
            If p > 0 Then
                If Not APS_Ativa(ops(p)) Then p = 0
            End If
            If p > 0 Then
                difHora = Abs(ops(p).Fim - ops(i).ini) > APS_EPS
                difDados = (Not APS_Ig(ops(p).maquina, ops(i).maquina)) Or _
                           (StrComp(ops(p).Produto, ops(i).Produto, vbTextCompare) <> 0) Or _
                           (StrComp(ops(p).lote, ops(i).lote, vbTextCompare) <> 0)
                If difHora Or difDados Then
                    nBk = nBk + 1
                    If Not APS_BackupLinha(APS_Aba(APS_ABA_OPS), ops(p).linha, backups(nBk)) Then
                        msg = msg & "Linha " & ops(p).linha & ": nao foi possivel fazer backup." & vbCrLf
                    End If
                End If
            End If
        End If
    Next i

    If Len(msg) > 0 Then
        MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel fazer backup de algumas linhas:" & vbCrLf & msg, vbExclamation, APS_TITULO
    End If

    ' REPARO: Setup/Limpeza gravado com ID provisorio "R<linha>" cuja producao ja recebeu ID definitivo
    ' (o vinculo se perdia). Religa pela linha, conferindo o produto.
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 And Len(ops(i).idPai) > 1 And Len(ops(i).idPai) <= 7 Then
            If UCase$(Left$(ops(i).idPai, 1)) = "R" And IsNumeric(mID$(ops(i).idPai, 2)) Then
                achou = False
                For j = 1 To n
                    If Len(ops(j).tipoAux) = 0 Then
                        If APS_Ig(ops(j).id, ops(i).idPai) Then achou = True: Exit For
                    End If
                Next j
                If Not achou Then
                    lin = CLng(mID$(ops(i).idPai, 2))
                    For j = 1 To n
                        If Len(ops(j).tipoAux) = 0 And ops(j).linha = lin Then
                            If StrComp(ops(j).Produto, ops(i).Produto, vbTextCompare) = 0 Then
                                ops(i).obs = Replace(ops(i).obs, ":" & ops(i).idPai & "]]", ":" & ops(j).id & "]]", 1, -1, vbTextCompare)
                                ops(i).idPai = ops(j).id
                                APS_GravarOp ops(i)
                                APS_SincronizarSetups = True
                            End If
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
    Next i

    For i = 1 To n
        If Len(ops(i).tipoAux) = 0 Then   ' i e producao
            p = APS_AcharAuxDaProducao(ops, n, ops(i).id, APS_TIPO_SETUP)
            If p = 0 Then p = APS_AcharAuxDaProducao(ops, n, ops(i).id, APS_TIPO_LIMPEZA)

            ' Setup/Limpeza nao tem status proprio e acompanha SEMPRE a producao (a producao e quem manda).
            ' So nao e tocado quando a producao esta cancelada (Setup inativo, herdado em APS_LerOps).
            If p > 0 Then
                If Not APS_Ativa(ops(p)) Then p = 0
            End If

            If p > 0 Then
                difHora = Abs(ops(p).Fim - ops(i).ini) > APS_EPS
                ' maquina/produto/lote do vinculo espelham a producao (como em APS_SalvarOperacaoComSetup)
                difDados = (Not APS_Ig(ops(p).maquina, ops(i).maquina)) Or _
                           (StrComp(ops(p).Produto, ops(i).Produto, vbTextCompare) <> 0) Or _
                           (StrComp(ops(p).lote, ops(i).lote, vbTextCompare) <> 0)

                If difHora Then
                    durPar = ops(p).dur
                    If durPar <= APS_EPS Then durPar = ops(p).Fim - ops(p).ini

                    fimPar = ops(i).ini
                    iniPar = APS_SubtrairTrabalho(fimPar, durPar)

                    If iniPar >= 0 Then
                        ops(p).ini = iniPar
                        ops(p).Fim = fimPar
                    Else
                        difHora = False
                    End If
                End If

                If difDados Then
                    ops(p).maquina = ops(i).maquina
                    ops(p).Produto = ops(i).Produto
                    ops(p).lote = ops(i).lote
                    APS_GravarOp ops(p)
                    APS_SincronizarSetups = True
                ElseIf difHora Then
                    APS_GravarHorario ops(p)
                    APS_SincronizarSetups = True
                End If
            End If
        End If
    Next i
    Exit Function

Falha:
    ' [PEN-B-21] Em caso de erro, restaura todas as linhas backupadas antes de sair.
    On Error Resume Next
    For iBk = 1 To nBk
        APS_RestaurarLinha APS_Aba(APS_ABA_OPS), backups(iBk)
    Next iBk
    MsgBox "Falha durante a sincroniza" & ChrW(231) & ChrW(227) & "o de setups: " & Err.Description & vbCrLf & _
           "Todas as linhas alteradas foram restauradas ao estado original.", vbCritical, APS_TITULO
    Err.Clear
    falhaRecuperou = True
End Function

' INICIO DO PROCESSO da producao k: o inicio do SETUP vinculado (quando ele precede a
' producao) ou, sem Setup valido antes dela, o inicio da propria producao.
Public Function APS_IniProcesso(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long) As Double
    Dim s As Long

    APS_IniProcesso = ops(k).ini
    s = APS_AcharAuxDaProducao(ops, n, ops(k).id, APS_TIPO_SETUP)
    If s > 0 Then
        If Not APS_Ig(ops(s).status, APS_ST_CANCELADA) Then
            If ops(s).Fim <= ops(k).ini + APS_EPS Then APS_IniProcesso = ops(s).ini
        End If
    End If
End Function

Public Function APS_EhGranulacao(ByVal maq As String) As Boolean
    APS_EhGranulacao = APS_Ig(maq, "Granula" & ChrW(231) & ChrW(227) & "o")
End Function

Public Function APS_CodMistura() As String
    APS_CodMistura = "906390"
End Function

Public Function APS_ProdutoPermitido(ByVal codigo As String, ByVal maq As String) As Boolean
    If APS_EhGranulacao(maq) Then
        APS_ProdutoPermitido = (Trim$(codigo) = APS_CodMistura())
    ElseIf APS_EhMediseal(maq) Then
        ' velocidade-base ja existente OU velocidade (caixas/min) cadastrada em Cadastros e Parametros
        APS_ProdutoPermitido = (APS_VelBaseDoCodigo(codigo) > 0) Or (APS_ParVel(maq, codigo) > 0)
    Else
        APS_ProdutoPermitido = (APS_IndiceProduto(codigo) > 0)
    End If
End Function

Public Function APS_MaquinasDoProduto(ByVal codigo As String, ByRef out() As String) As Long
    Dim m() As String, n As Long, i As Long, k As Long

    n = APS_Maquinas(m)
    ReDim out(1 To n)

    For i = 1 To n
        If APS_ProdutoPermitido(codigo, m(i)) Then
            k = k + 1
            out(k) = m(i)
        End If
    Next i

    APS_MaquinasDoProduto = k
End Function

'----------------------------------------------------------
' Coluna Setup de 02_Operacoes
'----------------------------------------------------------
' Valor da celula Setup -> dias (0 = sem Setup). Aceita tempo do Excel (1:30), texto "1:30" ou minutos (90).
Public Function APS_SetupDoValor(ByVal v As Variant) As Double
    Dim d As Double, s As String

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        If Len(s) = 0 Then Exit Function
        If IsNumeric(s) Then
            d = Val(Replace(s, ",", "."))
            If d >= 1# Then d = d / 1440#
        Else
            d = APS_ParaDuracao(s)
        End If
    ElseIf IsNumeric(v) Then
        d = CDbl(v)
        If d >= 1# Then d = d / 1440#          ' digitado em minutos (ex.: 90)
    End If

    If d > APS_EPS Then APS_SetupDoValor = APS_ArredMin(d)
End Function

' Rotulo do parametro cadastrado (APS_ParametrosSetup) cujo tempo e igual ao da coluna, para a maquina.
' Sem correspondencia: "Setup". tipo = LIMPEZA quando o rotulo e de limpeza; senao SETUP.
Public Function APS_RotuloPorMinutos(ByVal maq As String, ByVal dur As Double, ByRef tipo As String) As String
    Dim rot() As String, mi() As Long, n As Long, i As Long, m As Long

    tipo = APS_TIPO_SETUP
    n = APS_ParametrosSetup(maq, rot, mi)
    m = CLng(Round(dur * 1440#, 0))
    For i = 1 To n
        If mi(i) = m Then
            APS_RotuloPorMinutos = rot(i)
            If InStr(1, rot(i), "Limpeza", vbTextCompare) > 0 Then tipo = APS_TIPO_LIMPEZA
            Exit Function
        End If
    Next i
    APS_RotuloPorMinutos = "Setup"
End Function

' Monta (em memoria) o bloco de Setup/Limpeza da producao p a partir do tempo da coluna Setup.
' linha NEGATIVA = sem linha propria na planilha (APS_GravarOp/APS_GravarHorario/APS_LimparLinha tratam isso).
Public Sub APS_MontarAuxDaColuna(ByRef a As tOperacao, ByRef p As tOperacao, ByVal dur As Double)
    Dim tipo As String, rot As String

    rot = APS_RotuloPorMinutos(p.maquina, dur, tipo)

    a.linha = -p.linha
    a.id = "S" & p.id
    a.codigo = IIf(APS_Ig(tipo, APS_TIPO_LIMPEZA), APS_COD_LIMPEZA, APS_COD_SETUP)
    a.Produto = p.Produto
    a.lote = p.lote
    a.ordem = ""
    a.maquina = p.maquina
    a.caixas = 0
    a.oee = 0
    a.velBase = 0
    a.status = APS_ST_PLANEJADA          ' so em memoria (APS_LerOps herda o estado da producao)
    a.tipoAux = tipo
    a.idPai = p.id
    a.rotAux = rot
    a.obs = APS_MarcadorVinculo(tipo, p.id) & " " & rot & " - " & p.maquina
    a.dur = APS_ArredMin(dur)
    a.SemDur = False
    a.Fim = p.ini
    a.ini = APS_SubtrairTrabalho(a.Fim, a.dur)
    If a.ini < 0 Then a.ini = a.Fim - a.dur
End Sub

' Grava (ou limpa, dur <= 0) o tempo de Setup na linha da producao. A aba deve estar liberada pelo chamador.
Public Sub APS_GravarSetupDaOp(ByVal ws As Worksheet, ByVal linha As Long, ByVal durDias As Double)
    Dim c As Long

    If ws Is Nothing Then Exit Sub
    If linha < 2 Then Exit Sub
    c = APS_Col(ws, H_SET())
    If c = 0 Then Exit Sub

    If durDias > APS_EPS Then
        ws.Cells(linha, c).NumberFormat = "[h]:mm"
        ws.Cells(linha, c).Value2 = APS_ArredMin(durDias)
    Else
        ws.Cells(linha, c).ClearContents
    End If
    Central_MarcarLinha linha, "UPSERT"      ' FASE 6
End Sub

'----------------------------------------------------------
' Leitura de 02_Operacoes
' ID e OEE nao sao obrigatorios nem existem na planilha.
' O ID tecnico e mantido internamente pela linha da operacao.
'----------------------------------------------------------
Public Function APS_LerOps(ByRef ops() As tOperacao) As Long
    Dim ws As Worksheet, r As Long, ult As Long, n As Long
    Dim cCod As Long, cPro As Long, cLot As Long, cMaq As Long
    Dim cIni As Long, cFim As Long, cPla As Long
    Dim cObs As Long, cQtd As Long, cDur As Long, cId As Long
    Dim ai As Long, aj As Long
    Dim cOrd As Long, cOee As Long, cVel As Long, cQcp As Long
    Dim ini As Double, Fim As Double, dur As Double
    Dim x As Variant, mq As String
    Dim cSet As Long, setDur() As Double, nProd As Long
    Dim cAtr As Long, atr As Double, durCol As Double

    APS_OpsIgnoradas = 0
    ReDim ops(1 To 1)

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then
        MsgBox "A aba '" & APS_ABA_OPS & "' n" & ChrW(227) & "o foi encontrada.", _
               vbCritical, APS_TITULO
        APS_LerOps = -1
        Exit Function
    End If

    If APS_Col(ws, H_INI()) = 0 Then
        If Not APS_Migrar() Then
            APS_LerOps = -1
            Exit Function
        End If
    End If

    cCod = APS_Col(ws, H_COD())
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())
    cPla = APS_Col(ws, H_PLA())
    cObs = APS_Col(ws, H_OBS())
    cQtd = APS_Col(ws, H_QTD())
    cDur = APS_Col(ws, H_DUR())
    cOrd = APS_Col(ws, H_ORD())
    cOee = APS_Col(ws, H_OEE())
    cVel = APS_Col(ws, H_VEL())
    cQcp = APS_Col(ws, H_QCP())
    cSet = APS_Col(ws, H_SET())
    cAtr = APS_Col(ws, H_ATR())

    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cMaq = 0 Or _
       cIni = 0 Or cFim = 0 Or cPla = 0 Then

        MsgBox "A estrutura de 02_Operacoes n" & ChrW(227) & "o p" & _
               ChrW(244) & "de ser montada.", vbCritical, APS_TITULO

        APS_LerOps = -1
        Exit Function
    End If

    ' ID tecnico (coluna K): garante que toda linha de dados tenha ID persistente antes de ler.
    ' O ID vem da coluna K; a linha fisica NUNCA e a identidade da OP.
    APS_GarantirIDs
    cId = APS_Col(ws, H_ID())

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row

    If ult < 2 Then Exit Function

    ' o dobro: cada producao pode gerar, em memoria, um bloco de Setup (coluna Setup)
    ReDim ops(1 To (ult - 1) * 2)
    ReDim setDur(1 To ult - 1)

    For r = 2 To ult

        If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Or _
           Len(APS_Txt(ws.Cells(r, cCod).Value)) > 0 Then

            ini = APS_ParaDataHora(ws.Cells(r, cIni).Value)
            mq = APS_Txt(ws.Cells(r, cMaq).Value)

            If ini < 0 Or Len(mq) = 0 Then

                APS_OpsIgnoradas = APS_OpsIgnoradas + 1

            Else

                n = n + 1

                With ops(n)

                    .linha = r

                    ' identidade = ID da coluna K. "R<linha>" e so provisorio (K nao pode ser
                    ' gravada, ex.: aba protegida com senha); nunca e gravado em K.
                    .id = ""
                    If cId > 0 Then .id = APS_Txt(ws.Cells(r, cId).Value)
                    If Len(.id) = 0 Then .id = "R" & CStr(r)

                    .codigo = APS_Txt(ws.Cells(r, cCod).Value)
                    .Produto = APS_Txt(ws.Cells(r, cPro).Value)

                    If Len(.codigo) = 0 Then
                        .codigo = APS_CodigoDoNome(.Produto)
                    End If

                    .lote = APS_Txt(ws.Cells(r, cLot).Value)
                    If cOrd > 0 Then .ordem = APS_Txt(ws.Cells(r, cOrd).Value)
                    .maquina = mq

                    ' OEE (fracao) e velocidade-base GRAVADOS na operacao: a edicao os reutiliza
                    If cOee > 0 Then
                        x = ws.Cells(r, cOee).Value2
                        If Not IsEmpty(x) Then
                            If IsNumeric(x) Then
                                If CDbl(x) > 1# Then x = CDbl(x) / 100#   ' digitado como 70 em vez de 70%
                                If CDbl(x) > 0# And CDbl(x) <= 1# Then .oee = CDbl(x)
                            End If
                        End If
                    End If
                    If cVel > 0 Then
                        x = ws.Cells(r, cVel).Value2
                        If Not IsEmpty(x) Then
                            If IsNumeric(x) Then
                                If CDbl(x) > 0# Then .velBase = CDbl(x)
                            End If
                        End If
                    End If

                    .status = APS_Txt(ws.Cells(r, cPla).Value)

                    If Len(.status) = 0 Then
                        .status = APS_ST_PLANEJADA
                    End If

                    If cObs > 0 Then
                        .obs = APS_Txt(ws.Cells(r, cObs).Value)
                    End If

                    If cQtd > 0 Then
                        x = ws.Cells(r, cQtd).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                .caixas = CDbl(x)
                            End If
                        End If
                    End If

                    If cQcp > 0 Then
                        x = ws.Cells(r, cQcp).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                If CDbl(x) > 0# Then .comp = CDbl(x)
                            End If
                        End If
                    End If

                    dur = 0

                    If cDur > 0 Then
                        x = ws.Cells(r, cDur).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                dur = CDbl(x)
                            End If
                        End If
                    End If

                    durCol = dur                  ' duracao ORIGINAL gravada na coluna Duracao
                    atr = 0
                    If cAtr > 0 Then atr = APS_SetupDoValor(ws.Cells(r, cAtr).Value2)

                    Fim = APS_ParaDataHora(ws.Cells(r, cFim).Value)

                    If dur <= APS_EPS And Fim > ini Then
                        dur = Fim - ini
                    End If

                    dur = APS_ArredMin(dur)

                    ' o motor trabalha com o tempo TOTAL programado (original + atraso); a coluna Duracao segue original
                    If atr > APS_EPS And durCol > APS_EPS Then dur = APS_ArredMin(durCol + atr)
                    .atraso = atr

                    .ini = ini

                    If Fim > ini + APS_EPS Then

                        .Fim = Fim
                        .dur = dur
                        .SemDur = False

                        If .dur <= APS_EPS Then
                            .dur = APS_ArredMin(Fim - ini)
                        End If

                    ElseIf dur > APS_EPS Then

                        .dur = dur
                        .Fim = APS_ArredMin(ini + dur)
                        .SemDur = False

                    Else

                        .dur = 0
                        .SemDur = True
                        .Fim = ini + 1# / 24#

                    End If

                End With

                If cSet > 0 Then setDur(n) = APS_SetupDoValor(ws.Cells(r, cSet).Value2)

                APS_ClassificarAux ops(n)
            End If
        End If
    Next r

    ' [Setup na propria operacao] Cada producao com tempo na coluna Setup ganha, SO EM MEMORIA, o bloco de
    ' Setup/Limpeza que o motor e os cards ja conhecem (tipoAux/idPai/rotAux): termina no inicio da producao
    ' (Setup.fim = Producao.ini) e dura o tempo da coluna. Nao existe linha de Setup em 02_Operacoes.
    ' Linha de Setup do formato antigo ainda presente (idPai) prevalece ate APS_MigrarSetupsParaColuna recolhe-la.
    nProd = n
    For ai = 1 To nProd
        If Len(ops(ai).tipoAux) = 0 And setDur(ai) > APS_EPS Then
            If APS_AcharAuxDaProducao(ops, n, ops(ai).id, APS_TIPO_SETUP) = 0 And _
               APS_AcharAuxDaProducao(ops, n, ops(ai).id, APS_TIPO_LIMPEZA) = 0 Then
                n = n + 1
                APS_MontarAuxDaColuna ops(n), ops(ai), setDur(ai)
            End If
        End If
    Next ai

    ' [modelo novo] Setup/Limpeza NAO tem status proprio (celula vazia). Para o motor, o estado dele
    ' vem sempre da producao vinculada (idPai): producao cancelada = nao ocupa; em andamento ou
    ' concluida = ja executado (historico, nao e remanejado); senao = ativo. So em memoria.
    For ai = 1 To n
        If Len(ops(ai).tipoAux) > 0 And Len(Trim$(ops(ai).idPai)) > 0 Then
            ops(ai).status = APS_ST_PLANEJADA
            For aj = 1 To n
                If Len(ops(aj).tipoAux) = 0 Then
                    If APS_Ig(ops(aj).id, ops(ai).idPai) Then
                        If APS_Ig(ops(aj).status, APS_ST_CANCELADA) Then
                            ops(ai).status = APS_ST_CANCELADA
                        ElseIf APS_Ig(ops(aj).status, APS_ST_ANDAMENTO) Or APS_Ig(ops(aj).status, APS_ST_CONCLUIDA()) Then
                            ops(ai).status = APS_ST_CONCLUIDA()
                        End If
                        Exit For
                    End If
                End If
            Next aj
        End If
    Next ai

    APS_LerOps = n
End Function

'----------------------------------------------------------
' Relogio: a cada minuto confere se algum status mudou de faixa
' (Planejada > Em Andamento > Concluida) ou se 02_Operacoes foi alterada/excluida
'----------------------------------------------------------
Public Function APS_AssinaturaOps() As Double
    Dim ws As Worksheet, v As Variant, r As Long, c As Long, ult As Long, tot As Double, x As Variant, k As Long, s As String

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 2 Then Exit Function

    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, APS_NColDados(ws))).Value2
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            x = v(r, c)
            If Not IsError(x) Then
                If Not IsEmpty(x) Then
                    If VarType(x) = vbString Then
                        s = CStr(x)
                        k = Len(s)
                        If k > 0 Then tot = tot + k * (r * 11 + c) + AscW(Left$(s, 1)) * (r + c) + AscW(Right$(s, 1))
                    ElseIf IsNumeric(x) Then
                        tot = tot + CDbl(x) * (r * 7 + c)
                    End If
                End If
            End If
        Next c
    Next r
    APS_AssinaturaOps = tot
End Function

Public Function APS_LotesDoProduto(ByRef ops() As tOperacao, _
                                   ByVal n As Long, _
                                   ByVal codigo As String, _
                                   ByRef lotes() As String) As Long

    Dim i As Long, j As Long, k As Long
    Dim achou As Boolean

    ReDim lotes(1 To IIf(n < 1, 1, n))

    For i = 1 To n

        If Trim$(ops(i).codigo) = Trim$(codigo) And _
           Len(ops(i).lote) > 0 Then

            achou = False

            For j = 1 To k
                If APS_Ig(lotes(j), ops(i).lote) Then
                    achou = True
                    Exit For
                End If
            Next j

            If Not achou Then
                k = k + 1
                lotes(k) = ops(i).lote
            End If
        End If
    Next i

    APS_LotesDoProduto = k
End Function

Public Function APS_CodigoDoLote(ByRef ops() As tOperacao, _
                                 ByVal n As Long, _
                                 ByVal lote As String, _
                                 ByVal excluirID As String) As String

    Dim i As Long

    For i = 1 To n

        If APS_Ig(ops(i).lote, lote) And _
           Not APS_Ig(ops(i).id, excluirID) And _
           Not APS_EhAuto(ops(i)) Then

            APS_CodigoDoLote = ops(i).codigo
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------
' ID tecnico interno (coluna "ID" de 02_Operacoes = K, oculta)
' OP <> ID INTERNO <> LINHA: o ID mora na coluna K e nunca depende da posicao da linha.
' Mover, inserir ou apagar linhas nao muda o ID. O contador (nome oculto APS_CFG_PROXID,
' mesma tecnica das demais configuracoes) guarda o ultimo ID emitido: um ID emitido nunca
' e reutilizado, nem depois da exclusao da OP.
' O numero da linha aparece UMA UNICA vez, na MIGRACAO: linhas que ja existiam sem ID
' recebem como ID o numero que ja usavam (preserva vinculos [[APS_SETUP_DE:n]] e cards).
'----------------------------------------------------------

Public Function APS_ColID(ByVal ws As Worksheet, Optional ByVal criar As Boolean = False) As Long
    Dim c As Long

    c = APS_Col(ws, H_ID())

    If c = 0 And criar And Not Central_EstruturaTravada() Then
        c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

        With ws.Cells(1, c)
            .Value = H_ID()
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If

    APS_ColID = c
End Function

' Coluna "Quantidade (comprimidos)": criada no fim da tabela (nenhuma coluna existente e movida), so quando pedido.
Public Function APS_ColQComp(ByVal ws As Worksheet, Optional ByVal criar As Boolean = False) As Long
    Dim c As Long

    c = APS_Col(ws, H_QCP())

    If c = 0 And criar And Not Central_EstruturaTravada() Then
        c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

        With ws.Cells(1, c)
            .Value = H_QCP()
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        ws.Columns(c).ColumnWidth = 24
    End If

    APS_ColQComp = c
End Function

' Maior ID numerico gravado na coluna de ID
Private Function IdMaximo(ByVal ws As Worksheet, ByVal cId As Long) As Long
    Dim r As Long, ult As Long, v As Long

    If cId = 0 Then Exit Function

    ult = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row

    For r = 2 To ult
        v = CLng(Val(APS_Txt(ws.Cells(r, cId).Value)))
        If v > IdMaximo Then IdMaximo = v
    Next r
End Function

' Base do proximo ID = maior entre: contador gravado, maior ID existente e ultima linha usada
' (a ultima linha e so um PISO, para nunca colidir com IDs migrados da linha).
Private Function IdBase() As Long
    Dim ws As Worksheet, cId As Long, cPro As Long, u As Long, m As Long

    IdBase = CLng(Val(APS_CfgLer("PROXID", "0")))

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function

    cId = APS_Col(ws, H_ID())
    m = IdMaximo(ws, cId)
    If m > IdBase Then IdBase = m

    cPro = APS_Col(ws, H_PRO())
    If cPro > 0 Then
        u = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row
        If u > IdBase Then IdBase = u
    End If
End Function

' ID que a proxima OP nova vai receber (nao consome o contador)
Public Function APS_ProximoID() As String
    APS_ProximoID = CStr(IdBase() + 1)
End Function

' Emite um ID novo e avanca o contador
Public Function APS_NovoID() As String
    Dim v As Long

    ' FASE 6: em modo conectado o ApsId vem do bloco reservado no contador CENTRAL (nunca do contador local)
    If Central_Conectado() Then
        APS_NovoID = Central_ProximoIdCentral(True)
        If Len(APS_NovoID) > 0 Then APS_RegistrarID APS_NovoID
        Exit Function
    End If

    v = IdBase() + 1
    APS_CfgGravar "PROXID", CStr(v)

    ' confirma que o contrario foi gravado antes de entregar o ID
    If CLng(Val(APS_CfgLer("PROXID", "0"))) = v Then
        APS_NovoID = CStr(v)
    Else
        ' falha ao persistir o contrario: nao devemos entregar um ID cujo registro nao pode ser confirmado
        APS_NovoID = ""
    End If
End Function

' Garante que o contador nunca fique atras de um ID ja gravado
Public Sub APS_RegistrarID(ByVal id As String)
    Dim v As Long

    If Not IsNumeric(id) Then Exit Sub

    v = CLng(Val(id))
    If v > CLng(Val(APS_CfgLer("PROXID", "0"))) Then APS_CfgGravar "PROXID", CStr(v)
End Sub

' O ID ja esta gravado em outra linha (exceto a linha 'exceto')?
Public Function APS_IDEmUso(ByVal ws As Worksheet, ByVal cId As Long, ByVal id As String, _
                            ByVal exceto As Long) As Boolean
    Dim r As Long, ult As Long, v As String

    If cId = 0 Then Exit Function

    ult = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row

    For r = 2 To ult
        If r <> exceto Then
            v = APS_Txt(ws.Cells(r, cId).Value)
            If Len(v) > 0 Then
                If APS_Ig(v, id) Then
                    APS_IDEmUso = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' Toda linha de dados recebe ID persistente na coluna K (criada e oculta se preciso).
' Nao altera A:J. Linhas sem ID (ou com ID repetido) recebem ID novo do contador; na
' primeira execucao (sem contador), linhas sem ID recebem o numero que ja usavam.
' [modelo novo] Troca, nos marcadores [[APS_..._DE:id]] do Setup/Limpeza, o ID provisorio ("R<linha>")
' pelo ID definitivo da producao, para o vinculo (idPai) nao se perder quando o ID muda.
Private Sub APS_RemapearIdPai(ByVal ws As Worksheet, ByVal antigo As String, ByVal novo As String)
    Dim cObs As Long, r As Long, ult As Long, tx As String

    If Len(antigo) = 0 Or Len(novo) = 0 Then Exit Sub
    cObs = APS_Col(ws, H_OBS())
    If cObs = 0 Then Exit Sub
    ult = ws.Cells(ws.Rows.Count, cObs).End(xlUp).Row
    For r = 2 To ult
        tx = CStr(ws.Cells(r, cObs).Value)
        If InStr(1, tx, ":" & antigo & "]]", vbTextCompare) > 0 Then
            ws.Cells(r, cObs).Value = Replace(tx, ":" & antigo & "]]", ":" & novo & "]]", 1, -1, vbTextCompare)
        End If
    Next r
End Sub

Public Sub APS_GarantirIDs()
    If Central_EstruturaTravada() Then Exit Sub     ' FASE 6: modo conectado nao altera estrutura/IDs
    Dim ws As Worksheet, cId As Long, cPro As Long, cCod As Long
    Dim ult As Long, r As Long, i As Long, v As String, vistos As String
    Dim pend() As Long, np As Long
    Dim primeira As Boolean, prox As Long, mx As Long
    Dim est As Boolean, ocupou As Boolean, escondida As Boolean
    Dim ids() As String

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    If APS_Col(ws, H_INI()) = 0 Then Exit Sub       ' layout antigo: a migracao cuida disso

    cPro = APS_Col(ws, H_PRO())
    cCod = APS_Col(ws, H_COD())
    If cPro = 0 Or cCod = 0 Then Exit Sub

    cId = APS_Col(ws, H_ID())

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row
    r = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    If r > ult Then ult = r

    ReDim pend(1 To IIf(ult < 2, 1, ult))
    ReDim ids(1 To IIf(ult < 2, 1, ult))

    For r = 2 To ult
        If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Or Len(APS_Txt(ws.Cells(r, cCod).Value)) > 0 Then
            v = ""
            If cId > 0 Then v = APS_Txt(ws.Cells(r, cId).Value)

            If Len(v) > 0 And IsNumeric(v) And InStr(vistos, "|" & v & "|") = 0 Then
                vistos = vistos & "|" & v & "|"
            Else
                np = np + 1
                pend(np) = r
            End If
        End If
    Next r

    If cId > 0 Then escondida = CBool(ws.Columns(cId).Hidden)

    ' caminho rapido: nada a gravar e a coluna ja esta oculta
    If np = 0 And (cId = 0 Or escondida) Then Exit Sub

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    ocupou = True

    If cId = 0 Then cId = APS_ColID(ws, True)

    primeira = (Len(APS_CfgLer("PROXID", "")) = 0)
    prox = IdBase()

    For i = 1 To np
        r = pend(i)
        v = ""

        If primeira Then
            v = CStr(r)                                  ' migracao unica: mantem o numero que a linha ja usava
            If InStr(vistos, "|" & v & "|") > 0 Then v = ""
        End If

        If Len(v) = 0 Then
            prox = prox + 1
            v = CStr(prox)
            Do While InStr(vistos, "|" & v & "|") > 0
                prox = prox + 1
                v = CStr(prox)
            Loop
        End If

        vistos = vistos & "|" & v & "|"
        ws.Cells(r, cId).NumberFormat = "0"
        ws.Cells(r, cId).Value2 = CLng(v)
        ids(i) = v
        If CLng(v) > mx Then mx = CLng(v)
    Next i

    ' O ID provisorio "R<linha>" vira o ID permanente: o idPai gravado no Setup/Limpeza acompanha.
    For i = 1 To np
        APS_RemapearIdPai ws, "R" & pend(i), ids(i)
    Next i

    If mx > prox Then prox = mx
    If primeira Or prox > CLng(Val(APS_CfgLer("PROXID", "0"))) Then APS_CfgGravar "PROXID", CStr(prox)

    ws.Columns(cId).Hidden = True                        ' coluna tecnica

Sai:
    On Error Resume Next
    If ocupou Then APS_Ocupado = APS_Ocupado - 1
    If est Then APS_Reproteger ws, est
End Sub

' Limpa uma linha de dados (A:J e o ID da coluna K). A linha nao e excluida fisicamente.
Public Sub APS_LimparLinha(ByVal ws As Worksheet, ByVal linha As Long)
    Dim cId As Long, cOeeL As Long, cVelL As Long

    If linha < 2 Then Exit Sub

    Central_MarcarLinha linha, "EXCLUIR"      ' FASE 6: exclusao LOGICA central (le o ID antes de limpar); no-op em modo local

    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, APS_NColDados(ws))).ClearContents

    cId = APS_Col(ws, H_ID())
    If cId > 0 Then ws.Cells(linha, cId).ClearContents

    ' OEE e velocidade-base fazem parte da operacao: saem junto
    cOeeL = APS_Col(ws, H_OEE())
    If cOeeL > 0 Then ws.Cells(linha, cOeeL).ClearContents
    cVelL = APS_Col(ws, H_VEL())
    If cVelL > 0 Then ws.Cells(linha, cVelL).ClearContents
End Sub

' Grava so a identidade (Codigo, Produto, Lote, Maquina) de uma linha ja existente
Public Sub APS_GravarIdentidade(ByRef o As tOperacao)
    Dim ws As Worksheet, cCod As Long, cPro As Long, cLot As Long, cMaq As Long

    If o.linha < 2 Then Exit Sub

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub

    cCod = APS_Col(ws, H_COD())
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())

    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cMaq = 0 Then Exit Sub

    ws.Cells(o.linha, cCod).NumberFormat = "@"
    ws.Cells(o.linha, cCod).Value = o.codigo

    ws.Cells(o.linha, cPro).Value = o.Produto

    ws.Cells(o.linha, cLot).NumberFormat = "@"
    ws.Cells(o.linha, cLot).Value = o.lote

    ws.Cells(o.linha, cMaq).Value = o.maquina
End Sub

'----------------------------------------------------------
' Gravacao em 02_Operacoes
'----------------------------------------------------------
Public Sub APS_GravarOp(ByRef o As tOperacao)

    Dim ws As Worksheet, r As Long

    Dim cCod As Long, cPro As Long, cLot As Long
    Dim cMaq As Long, cIni As Long, cFim As Long
    Dim cPla As Long, cObs As Long, cQtd As Long, cDur As Long
    Dim cId As Long, idAnt As String
    Dim cOrd As Long, cOee As Long, cVel As Long, cAtr As Long, cQcp As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then Exit Sub

    ' Setup/Limpeza NAO tem linha propria: gravar = registrar o tempo na coluna Setup da producao
    If o.linha < 0 Then
        APS_GravarSetupDaOp ws, -o.linha, o.dur
        Exit Sub
    End If

    cCod = APS_Col(ws, H_COD())
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())
    cPla = APS_Col(ws, H_PLA())
    cObs = APS_Col(ws, H_OBS())
    cQtd = APS_Col(ws, H_QTD())
    cDur = APS_Col(ws, H_DUR())
    cOrd = APS_Col(ws, H_ORD())
    cOee = APS_Col(ws, H_OEE())
    cVel = APS_Col(ws, H_VEL())
    cAtr = APS_Col(ws, H_ATR())
    cQcp = APS_ColQComp(ws, o.comp > 0)     ' so cria a coluna quando ha quantidade em comprimidos para gravar

    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cMaq = 0 Or _
       cIni = 0 Or cFim = 0 Or cPla = 0 Then Exit Sub

    If o.linha > 0 Then

        r = o.linha

    Else

        r = APS_LinhaLivre(ws, cCod, cPro, cLot)

        o.linha = r

    End If

    ' ID tecnico persistente (coluna K). Nunca derivado da linha nem reutilizado.
    cId = APS_ColID(ws, True)
    If cId > 0 Then
        If Len(o.id) = 0 Or Not IsNumeric(o.id) Then
            idAnt = o.id
            o.id = APS_NovoID()
            If Len(o.id) = 0 And Central_Conectado() Then Exit Sub      ' FASE 6: sem ApsId central nada e gravado
            If Len(idAnt) > 0 Then APS_RemapearIdPai ws, idAnt, o.id   ' o Setup/Limpeza vinculado acompanha o novo ID
        ElseIf APS_IDEmUso(ws, cId, o.id, r) Then
            o.id = APS_NovoID()
            If Len(o.id) = 0 And Central_Conectado() Then Exit Sub
        Else
            APS_RegistrarID o.id
        End If
        ws.Cells(r, cId).NumberFormat = "0"
        ws.Cells(r, cId).Value2 = CLng(Val(o.id))
    End If

    ' cabecalho azul, dados normais: a linha de dados recebe formato proprio
    APS_FormatarLinhaDados ws, r

    ws.Cells(r, cCod).NumberFormat = "@"
    ws.Cells(r, cCod).Value = o.codigo

    ws.Cells(r, cPro).Value = o.Produto

    ws.Cells(r, cLot).NumberFormat = "@"
    ws.Cells(r, cLot).Value = o.lote

    ws.Cells(r, cMaq).Value = o.maquina

    ws.Cells(r, cIni).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(r, cIni).Value2 = o.ini

    ws.Cells(r, cFim).NumberFormat = "dd/mm/yyyy hh:mm"

    If o.SemDur Then
        ws.Cells(r, cFim).ClearContents
    Else
        ws.Cells(r, cFim).Value2 = o.Fim
    End If

    ' Setup/Limpeza nao tem status proprio: a celula fica vazia (o estado vem da producao, ver APS_LerOps)
    If Len(o.tipoAux) > 0 Then
        ws.Cells(r, cPla).ClearContents
    Else
        ws.Cells(r, cPla).Value = o.status
    End If
    If cObs > 0 Then ws.Cells(r, cObs).Value = o.obs

    ' OP, OEE e velocidade-base: gravados com a operacao (vazio limpa a celula)
    If cOrd > 0 Then
        ws.Cells(r, cOrd).NumberFormat = "@"
        If Len(o.ordem) > 0 Then ws.Cells(r, cOrd).Value = o.ordem Else ws.Cells(r, cOrd).ClearContents
    End If
    If cOee > 0 Then
        If o.oee > 0 Then
            ws.Cells(r, cOee).NumberFormat = "0.0%"
            ws.Cells(r, cOee).Value2 = o.oee
        Else
            ws.Cells(r, cOee).ClearContents
        End If
    End If
    If cVel > 0 Then
        If o.velBase > 0 Then
            ws.Cells(r, cVel).NumberFormat = "General"
            ws.Cells(r, cVel).Value2 = o.velBase
        Else
            ws.Cells(r, cVel).ClearContents
        End If
    End If

    If cQtd > 0 Then
        If o.caixas > 0 Then
            ws.Cells(r, cQtd).NumberFormat = "General"
            ws.Cells(r, cQtd).Value2 = o.caixas
        Else
            ws.Cells(r, cQtd).ClearContents
        End If
    End If

    ' Quantidade OFICIAL em comprimidos (a coluna de caixas acima e a quantidade derivada, usada pelos demais modulos)
    If cQcp > 0 Then
        If o.comp > 0 Then
            ws.Cells(r, cQcp).NumberFormat = "#,##0"
            ws.Cells(r, cQcp).Value2 = o.comp
        Else
            ws.Cells(r, cQcp).ClearContents
        End If
    End If

    If cDur > 0 Then
        ws.Cells(r, cDur).NumberFormat = "[h]:mm"

        If o.SemDur Then
            ws.Cells(r, cDur).ClearContents
        Else
            ws.Cells(r, cDur).Value2 = APS_DurOriginal(o)      ' Duracao = ORIGINAL (o atraso vai na coluna Atraso)
        End If
    End If

    If cAtr > 0 And Len(o.tipoAux) = 0 Then
        If o.atraso > APS_EPS And Not o.SemDur Then
            ws.Cells(r, cAtr).NumberFormat = "[h]:mm"
            ws.Cells(r, cAtr).Value2 = APS_ArredMin(o.atraso)
        Else
            ws.Cells(r, cAtr).ClearContents
        End If
    End If

    Central_MarcarLinha r, "UPSERT"      ' FASE 6: pendente de envio (no-op em modo local)

End Sub

'----------------------------------------------------------
' Salvar (nova ou editada)
'----------------------------------------------------------
' setupDias: -1 = nao mexe no Setup; 0 = remove o Setup; > 0 = tempo do Setup (dias). Producao e Setup entram no MESMO
' encaixe e na MESMA gravacao: ou os dois sao gravados, ou nada e gravado.
Public Function APS_SalvarOperacao(ByRef nv As tOperacao, ByVal ehNovo As Boolean, _
                                   Optional ByVal setupDias As Double = -1) As Boolean
    Dim pa As Long, mudouSetup As Boolean
    Dim ops() As tOperacao, base() As tOperacao, mudou() As Boolean, conf() As Boolean
    Dim n As Long, k As Long, i As Long, qtd As Long, calOk As Boolean
    Dim trav As Boolean, manual As Boolean, mudouHorario As Boolean
    Dim txt As String, lista As String, pedido As Double
    Dim wsO As Worksheet, estO As Boolean
    Dim temAnt As Boolean, antIni As Double, antFim As Double, antMaq As String, pkk As Long

    On Error GoTo Falha

    If ehNovo Then If Not Central_PodeCriar() Then Exit Function     ' FASE 6: sem ApsId central nao cria

    APS_GarantirIDs
    n = APS_LerOps(ops)
    If n < 0 Then Exit Function
    ReDim Preserve ops(1 To n + 2)

    For i = 1 To n
        If APS_Ig(ops(i).id, nv.id) Then k = i: Exit For
    Next i

    If k = 0 Then
        If Not ehNovo Then
            MsgBox "A opera" & ChrW(231) & ChrW(227) & "o n" & ChrW(227) & "o foi encontrada em 02_Operacoes.", vbExclamation, APS_TITULO
            Exit Function
        End If
        n = n + 1
        k = n
        nv.linha = 0
    Else
        nv.linha = ops(k).linha
        temAnt = True: antIni = ops(k).ini: antFim = ops(k).Fim: antMaq = ops(k).maquina
        ' status escolhido a mao no formulario: o automatico nao sobrescreve de imediato
        manual = Not APS_Ig(nv.status, ops(k).status)
        mudouHorario = (Abs(nv.ini - ops(k).ini) > APS_EPS) Or (Abs(nv.dur - ops(k).dur) > APS_EPS) Or _
                       (Not APS_Ig(nv.maquina, ops(k).maquina))
    End If
    If ehNovo Then mudouHorario = True
    ops(k) = nv
    pedido = nv.ini

    ' Setup pedido entra no encaixe como bloco logico com a producao (ainda em memoria)
    If setupDias >= 0 And APS_Ativa(ops(k)) Then
        pa = APS_AcharAuxDaProducao(ops, n, nv.id, APS_TIPO_SETUP)
        If pa = 0 Then pa = APS_AcharAuxDaProducao(ops, n, nv.id, APS_TIPO_LIMPEZA)
        If setupDias > APS_EPS Then
            If pa > 0 Then
                If Abs(ops(pa).dur - setupDias) > APS_EPS Then mudouSetup = True
                ops(pa).dur = setupDias
                ops(pa).SemDur = False
            Else
                n = n + 1
                APS_MontarAuxDaColuna ops(n), ops(k), setupDias
                mudouSetup = True
            End If
        ElseIf pa > 0 Then
            ops(pa).status = APS_ST_CANCELADA      ' sai do encaixe; a coluna e limpa na gravacao
            mudouSetup = True
        End If
    End If

    ReDim base(1 To n)
    For i = 1 To n
        base(i) = ops(i)
    Next i

    ' Posicao escolhida pelo usuario (nova OP, ou inicio/duracao/maquina alterados): a operacao fica onde foi
    ' colocada e as demais da maquina andam em cascata (ver APS_Encaixar, parametro prioridade).
    ' Antiga posicao do Setup: base(pkk) guarda o Setup como estava antes de mover a producao.
    If temAnt Then
        pkk = APS_AcharAuxDaProducao(base, n, ops(k).id, APS_TIPO_SETUP)
        If pkk = 0 Then pkk = APS_AcharAuxDaProducao(base, n, ops(k).id, APS_TIPO_LIMPEZA)
        If pkk > 0 Then
            If base(pkk).ini < antIni Then antIni = base(pkk).ini
        End If
    End If

    qtd = APS_Encaixar(ops, n, k, mudou, calOk, False, Not (mudouHorario Or mudouSetup), mudouHorario)
    If Not calOk Then
        If Len(APS_MotivoEncaixe) > 0 Then
            MsgBox APS_MotivoEncaixe, vbExclamation, APS_TITULO
        ElseIf APS_ForaDoMes Then
            MsgBox APS_MsgForaDoMes(), vbExclamation, APS_TITULO
        Else
            MsgBox "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO.", vbExclamation, APS_TITULO
        End If
        Exit Function
    End If

    Set wsO = APS_Aba(APS_ABA_OPS)
    estO = APS_Liberar(wsO)
    APS_Ocupado = APS_Ocupado + 1
    trav = True
    APS_GravarOp ops(k)
    nv.id = ops(k).id          ' devolve ao chamador a identidade FINAL gravada (ID e linha)
    nv.linha = ops(k).linha
    If setupDias >= 0 And APS_Ativa(ops(k)) Then
        APS_GravarSetupDaOp wsO, ops(k).linha, setupDias   ' tempo do Setup na MESMA linha (0 = limpa)
    End If
    For i = 1 To n
        If i <> k Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i
    If setupDias >= 0 And APS_Ativa(ops(k)) Then APS_RemoverAuxDaProducao ops, n, ops(k).id, 0   ' linha de Setup antiga, se existir
    If manual Then APS_MapaMarcar ops(k).linha, ops(k).ini, ops(k).Fim
    ' operacao existente movida para um horario POSTERIOR sem encostar no antigo: o espaco que ela deixou
    ' e um vao aberto por evento e pode ser fechado (as seguintes voltam), como na exclusao.
    If temAnt And mudouHorario Then
        If APS_Ig(antMaq, ops(k).maquina) And ops(k).ini > antFim + APS_EPS Then
            APS_VaoPendente antMaq, antIni, antFim
        End If
    End If
    APS_Ocupado = APS_Ocupado - 1
    trav = False
    APS_Reproteger wsO, estO
    estO = False

    APS_Atualizar True

    ' resumo do que o sistema ajustou
    If Abs(ops(k).ini - pedido) > APS_EPS Then
        txt = "In" & ChrW(237) & "cio ajustado (calend" & ChrW(225) & "rio / m" & ChrW(225) & "quina ocupada): " & _
              Format$(pedido, "dd/mm hh:nn") & " > " & Format$(ops(k).ini, "dd/mm hh:nn") & vbCrLf
    End If
    If qtd > 0 Then
        txt = txt & qtd & " opera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & "es) seguinte(s) deslocada(s) em cascata:" & vbCrLf
        For i = 1 To n
            If mudou(i) Then
                txt = txt & "  " & ops(i).Produto & " " & ops(i).lote & ": " & Format$(base(i).ini, "dd/mm hh:nn") & " > " & Format$(ops(i).ini, "dd/mm hh:nn") & _
                      " (fim " & Format$(ops(i).Fim, "dd/mm hh:nn") & ")" & vbCrLf
            End If
        Next i
    End If
    APS_DetectarConflitos ops, n, conf
    lista = APS_ListaConflitos(ops, n, k)
    If Len(lista) > 0 Then
        txt = txt & vbCrLf & "ATEN" & ChrW(199) & ChrW(195) & "O: conflito com opera" & ChrW(231) & ChrW(227) & "o que n" & ChrW(227) & "o pode ser movida: " & lista
    End If
    If Len(txt) > 0 Then MsgBox txt, IIf(Len(lista) > 0, vbExclamation, vbInformation), APS_TITULO

    APS_SalvarOperacao = True
    Exit Function

Falha:
    On Error Resume Next
    If trav Then APS_Ocupado = APS_Ocupado - 1
    If Not wsO Is Nothing Then APS_Reproteger wsO, estO
    MsgBox "Erro ao salvar a opera" & ChrW(231) & ChrW(227) & "o: " & Err.Description, vbCritical, APS_TITULO
End Function

'----------------------------------------------------------
' APS_SalvarOperacaoComSetup - RECONSTRUCAO COMPATIVEL - NAO CODIGO ORIGINAL.
' O Setup/Limpeza escolhido no formulario NAO vira uma linha: o tempo e gravado na coluna Setup da MESMA linha da
' producao. Producao e Setup sao encaixados juntos (APS_SalvarOperacao, parametro setupDias) e gravados na mesma
' passada: se nao couberem, NADA e gravado e o motivo e informado. Quem escolhe o parametro e o usuario.
Public Function APS_SalvarOperacaoComSetup(ByRef o As tOperacao, ByVal mNovo As Boolean, _
                                            ByVal temSetup As Boolean, ByVal setupRotulo As String, _
                                            ByVal setupMin As Long, ByRef msg As String) As Boolean
    Dim setupDias As Double, ws As Worksheet

    On Error GoTo Falha

    msg = ""

    If Len(Trim$(o.id)) = 0 Then
        msg = "A opera" & ChrW(231) & ChrW(227) & "o precisa ter ID para vincular Setup/Limpeza."
        Exit Function
    End If

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function
    If APS_Col(ws, H_SET()) = 0 Then
        msg = "A coluna Setup n" & ChrW(227) & "o existe em 02_Operacoes; nada foi gravado."
        Exit Function
    End If

    ' -1 = nao mexe no Setup (producao cancelada ou tempo nao informado); 0 = sem Setup; > 0 = tempo
    setupDias = -1
    If APS_Ativa(o) Then
        If Not temSetup Then
            setupDias = 0
        ElseIf setupMin > 0 Then
            setupDias = setupMin / 1440
            If APS_SubtrairTrabalho(o.ini, setupDias) < 0 Then
                msg = "N" & ChrW(227) & "o h" & ChrW(225) & " calend" & ChrW(225) & "rio de trabalho suficiente para o Setup/Limpeza. Nada foi gravado."
                Exit Function
            End If
        End If
    End If

    APS_SalvarOperacaoComSetup = APS_SalvarOperacao(o, mNovo, setupDias)
    Exit Function

Falha:
    msg = "Erro ao salvar a opera" & ChrW(231) & ChrW(227) & "o: " & Err.Description
End Function

'----------------------------------------------------------
' Primeira linha de dados livre (reaproveita linhas esvaziadas por exclusao)
'----------------------------------------------------------
Public Function APS_LinhaLivre(ByVal ws As Worksheet, ByVal cCod As Long, ByVal cPro As Long, _
                               ByVal cLot As Long) As Long
    Dim r As Long, ult As Long

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row > ult Then ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row

    For r = 2 To ult
        If Len(APS_Txt(ws.Cells(r, cPro).Value)) = 0 And _
           Len(APS_Txt(ws.Cells(r, cCod).Value)) = 0 And _
           Len(APS_Txt(ws.Cells(r, cLot).Value)) = 0 Then
            APS_LinhaLivre = r
            Exit Function
        End If
    Next r

    APS_LinhaLivre = ult + 1
    If APS_LinhaLivre < 2 Then APS_LinhaLivre = 2
End Function

'----------------------------------------------------------
' Formato de uma linha de dados de 02_Operacoes: sem preenchimento,
' fonte automatica, sem negrito. O azul e SO do cabecalho (linha 1).
'----------------------------------------------------------
Public Sub APS_FormatarLinhaDados(ByVal ws As Worksheet, ByVal r As Long)
    If r < 2 Then Exit Sub
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, APS_NColDados(ws)))
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Font.Bold = False
        .VerticalAlignment = xlCenter
    End With
End Sub

'----------------------------------------------------------
' MES DE PLANEJAMENTO = celula 1_Planejamento!A1 (data real; exibe "Mes: setembro/2026").
' Vazia/texto = sem definicao = sem limite mensal.
' Devolve o inicio do mes (inclusive) e o dia 1 do mes seguinte (exclusivo).
' O botao Organizar Automaticamente (aba SAP) deve chamar esta mesma funcao / APS_Encaixar.
'----------------------------------------------------------
Public Function APS_LimiteMes(ByRef mesIni As Double, ByRef mesFim As Double) As Boolean
    Dim ws As Worksheet, v As Variant, d As Date

    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Function
    v = ws.Range("A1").Value2
    If IsError(v) Then Exit Function
    If VarType(v) <> vbDouble Then Exit Function
    If v < 36526 Then Exit Function

    d = CDate(Int(v))
    mesIni = CDbl(DateSerial(Year(d), Month(d), 1))
    mesFim = CDbl(DateSerial(Year(d), Month(d) + 1, 1))
    APS_LimiteMes = True
End Function

' Protecao mensal: APS_ForaDoMes / APS_MesRef sao declaradas na secao de declaracoes do topo do modulo.

Public Function APS_MsgForaDoMes() As String
    APS_MsgForaDoMes = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel encaixar a opera" & ChrW(231) & ChrW(227) & "o dentro do m" & ChrW(234) & "s do planejamento (" & _
        Format$(APS_MesRef, "mm/yyyy") & "). O planejamento n" & ChrW(227) & "o altera outros meses." & _
        vbCrLf & "Nenhuma opera" & ChrW(231) & ChrW(227) & "o foi alterada. (M" & ChrW(234) & "s definido em 1_Planejamento, c" & ChrW(233) & "lula A1.)"
End Function

'----------------------------------------------------------
' Mediseal: comprimidos = caixas x 30; velocidade efetiva = base x OEE;
' minutos = comprimidos / velocidade efetiva
'----------------------------------------------------------
Public Function APS_CalcMediseal(ByVal caixas As Double, ByVal velBase As Double, ByVal oee As Double, _
                                 ByRef comprimidos As Double, ByRef velEf As Double, _
                                 ByRef minutos As Double) As Boolean
    If caixas <= 0 Or velBase <= 0 Or oee <= 0 Then Exit Function
    comprimidos = caixas * APS_CP_CAIXA
    velEf = velBase * oee
    minutos = Round(comprimidos / velEf, 0)
    APS_CalcMediseal = (minutos > 0)
End Function


Public Sub APS_PrepararMes()
    Dim ws As Worksheet, est As Boolean, fmt As String

    On Error GoTo Sai
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Sub
    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1

    fmt = """       PLANEJAMENTO """ & vbLf & """M" & ChrW(234) & "s: ""mmmm/yyyy"
    With ws.Range("A1")
        .NumberFormat = fmt
        .WrapText = True
        .VerticalAlignment = xlTop
        If .Font.Size > 11 Then .Font.Size = 11
        .Validation.Delete
        .Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlGreaterEqual, Formula1:="=DATE(2020,1,1)"
        .Validation.IgnoreBlank = True
        .Validation.InputTitle = "M" & ChrW(234) & "s do Planejamento"
        .Validation.InputMessage = "Digite o m" & ChrW(234) & "s, ex.: 01/09/2026. O motor de cascata s" & ChrW(243) & " altera opera" & ChrW(231) & ChrW(245) & "es deste m" & ChrW(234) & "s."
        .Validation.ErrorTitle = "M" & ChrW(234) & "s do Planejamento"
        .Validation.ErrorMessage = "Informe uma data v" & ChrW(225) & "lida (ex.: 01/09/2026)."
    End With

Sai:
    On Error Resume Next
    If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    If est Then APS_Reproteger ws, est
End Sub

Public Sub APS_CorrigirFormatoDados()
    Dim ws As Worksheet, est As Boolean, r As Long, ult As Long, c As Long, ruim As Boolean

    On Error GoTo Sai
    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub

    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 2 Then Exit Sub

    For r = 2 To ult
        ruim = False
        For c = 1 To APS_NColDados(ws)      ' todas as colunas de dados (inclui OP, Setup e Atraso)
            If ws.Cells(r, c).Interior.Color = APS_COR_CAB Then ruim = True: Exit For
        Next c
        If ruim Then
            If Not est Then est = APS_Liberar(ws)
            APS_FormatarLinhaDados ws, r
        End If
    Next r

Sai:
    On Error Resume Next
    If est Then APS_Reproteger ws, est
End Sub

Public Sub APS_GravarHorario(ByRef o As tOperacao)

    Dim ws As Worksheet, cIni As Long, cFim As Long, cDur As Long

    If o.linha < 2 Then Exit Sub

    Set ws = APS_Aba(APS_ABA_OPS)

    cIni = APS_Col(ws, H_INI())
    cFim = APS_Col(ws, H_FIM())
    cDur = APS_Col(ws, H_DUR())

    If cIni = 0 Or cFim = 0 Then Exit Sub

    ws.Cells(o.linha, cIni).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(o.linha, cIni).Value2 = o.ini

    If Not o.SemDur Then
        ws.Cells(o.linha, cFim).NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Cells(o.linha, cFim).Value2 = o.Fim

        ' A duracao acompanha o fim: e a MESMA o.dur (tempo de trabalho pelo calendario) que
        ' ja define o fim. Nao ha segunda regra de duracao.
        If cDur > 0 And o.dur > APS_EPS Then
            ws.Cells(o.linha, cDur).NumberFormat = "[h]:mm"
            ws.Cells(o.linha, cDur).Value2 = APS_DurOriginal(o)
        End If
    End If
    Central_MarcarLinha o.linha, "UPSERT"      ' FASE 6
End Sub

Public Sub APS_GravarStatus(ByVal linha As Long, ByVal status As String)

    Dim ws As Worksheet, c As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    c = APS_Col(ws, H_PLA())

    If c > 0 And linha >= 2 Then
        ws.Cells(linha, c).Value = status
        Central_MarcarLinha linha, "STATUSAUTO"      ' FASE 6: se houver 412, a alteracao automatica e descartada
    End If
End Sub

'----------------------------------------------------------
' EXCLUSAO DE OPERACAO
' Remove os dados da operacao da 02_Operacoes sem excluir
' fisicamente a linha, preservando os IDs das demais OPs.
' Depois atualiza o planejamento e os cards.
'----------------------------------------------------------
Public Function APS_ExcluirOperacao(ByVal id As String) As Boolean

    Dim ws As Worksheet
    Dim ops() As tOperacao
    Dim n As Long
    Dim i As Long
    Dim linha As Long
    Dim est As Boolean

    On Error GoTo Falha

    If Len(Trim$(id)) = 0 Then Exit Function

    n = APS_LerOps(ops)

    If n < 0 Then Exit Function

    For i = 1 To n
        If APS_Ig(ops(i).id, id) Then
            linha = ops(i).linha
            Exit For
        End If
    Next i

    If linha <= 0 Then
        MsgBox "Opera" & ChrW(231) & ChrW(227) & "o " & id & _
               " n" & ChrW(227) & "o encontrada.", _
               vbExclamation, APS_TITULO
        Exit Function
    End If

    ' linhas Setup/Limpeza vinculadas a esta OP: saem junto (nenhum registro orfao)
    Dim auxLinhas() As Long, nAux As Long
    ReDim auxLinhas(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 Then
            If APS_Ig(ops(i).idPai, id) Then
                nAux = nAux + 1
                auxLinhas(nAux) = ops(i).linha
            End If
        End If
    Next i

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then
        MsgBox "A aba " & APS_ABA_OPS & _
               " n" & ChrW(227) & "o foi encontrada.", _
               vbCritical, APS_TITULO
        Exit Function
    End If

    est = APS_Liberar(ws)

    ' A:J + ID (K) da linha sao limpos; a linha nao e excluida fisicamente. O ID excluido
    ' NAO e reutilizado (o contador so avanca). Remove antes as linhas Setup/Limpeza vinculadas.
    APS_Ocupado = APS_Ocupado + 1
    For i = 1 To nAux
        APS_LimparLinha ws, auxLinhas(i)
    Next i
    APS_LimparLinha ws, linha
    APS_Ocupado = APS_Ocupado - 1

    APS_Reproteger ws, est

    ' Reconstroi os cards e atualiza conflitos/status.
    APS_Atualizar True

    APS_ExcluirOperacao = True

    MsgBox "Opera" & ChrW(231) & ChrW(227) & "o " & id & _
           " exclu" & ChrW(237) & "da com sucesso.", _
           vbInformation, APS_TITULO

    Exit Function

Falha:

    On Error Resume Next

    ' [CORRECAO] O APS_Ocupado foi adquirido antes da limpeza (linha 2419).
    ' Em caso de erro, restaura o counter e a protecao da planilha antes de sair.
    APS_Ocupado = APS_Ocupado - 1
    If Not ws Is Nothing Then
        APS_Reproteger ws, est
    End If

    MsgBox "Erro ao excluir a opera" & ChrW(231) & ChrW(227) & "o: " & _
           Err.Description, vbCritical, APS_TITULO

End Function

' [ADAPTACAO modelo novo] True se existe, em 02_Operacoes, uma linha de DADOS cujo ID (coluna K)
' e idPai. Protege contra remover um Setup/Limpeza cuja producao existe mas foi ignorada por
' APS_LerOps (ex.: data/maquina apagada durante uma edicao). Sem como confirmar -> True (nao remove).
' [CORRECAO PEN-B-05] Verifica se o idPai ainda aponta para uma producao real na planilha.
' Diferenciamos tres situacoes:
'   - productions real encontrada (tipoAux = 0 e com dados) -> nao e orfao
'   - linha encontrada mas e Setup/Limpeza (tipoAux > 0) -> ambiguo, nao tratar como pai
'   - nenhuma linha encontrada -> orfao
' O padrao inicial e False (nao assume que o pai existe). A funcao so retorna True quando
' encontra uma linha que claramente e uma producao.
Private Function APS_PaiNaPlanilha(ByVal idPai As String) As Boolean
    Dim ws As Worksheet, cId As Long, cPro As Long, cCod As Long, cObs As Long, r As Long, ult As Long
    Dim txObs As String

    APS_PaiNaPlanilha = False
    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function
    cId = APS_Col(ws, H_ID())
    cPro = APS_Col(ws, H_PRO())
    cCod = APS_Col(ws, H_COD())
    cObs = APS_Col(ws, H_OBS())
    If cId = 0 Or (cPro = 0 And cCod = 0) Then Exit Function

    ' ID provisorio "R<linha>": a producao existe se aquela linha tem dados de producao
    ' (e nao e uma linha de Setup/Limpeza, que teria marcador nas Observacoes).
    If UCase$(Left$(idPai, 1)) = "R" And IsNumeric(mID$(idPai, 2)) And Len(idPai) <= 7 Then
        r = CLng(mID$(idPai, 2))
        If r >= 2 Then
            If cObs > 0 Then
                txObs = APS_Txt(ws.Cells(r, cObs).Value)
                If InStr(1, txObs, "[[APS_", vbTextCompare) > 0 Then Exit Function   ' linha e Setup/Limpeza, nao e pai
            End If
            If cPro > 0 Then
                If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Then APS_PaiNaPlanilha = True: Exit Function
            End If
            If cCod > 0 Then
                If Len(APS_Txt(ws.Cells(r, cCod).Value)) > 0 Then APS_PaiNaPlanilha = True: Exit Function
            End If
        End If
        Exit Function
    End If

    ult = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row
    For r = 2 To ult
        If APS_Ig(APS_Txt(ws.Cells(r, cId).Value), idPai) Then
            ' confirma que esta linha e de fato uma producao (tem dados e nao e Setup/Limpeza)
            If cObs > 0 Then
                txObs = APS_Txt(ws.Cells(r, cObs).Value)
                If InStr(1, txObs, "[[APS_", vbTextCompare) > 0 Then GoTo Proximo   ' linha encontrada mas e Setup/Limpeza: nao e pai
            End If
            If cPro > 0 Then
                If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Then APS_PaiNaPlanilha = True: Exit Function
            End If
            If cCod > 0 Then
                If Len(APS_Txt(ws.Cells(r, cCod).Value)) > 0 Then APS_PaiNaPlanilha = True: Exit Function
            End If
        End If
Proximo:
    Next r
End Function


' [ADAPTACAO modelo novo] Remove as linhas Setup/Limpeza vinculadas a idProd, exceto a de
' indice "manter" (0 = remove todas). PREPARACAO = uma unica linha vinculada por producao.
' Devolve quantas foram removidas.
Private Function APS_RemoverAuxDaProducao(ByRef ops() As tOperacao, ByVal n As Long, _
                                          ByVal idProd As String, ByVal manter As Long) As Long
    Dim ws As Worksheet, est As Boolean, i As Long

    If Len(Trim$(idProd)) = 0 Then Exit Function
    For i = 1 To n
        If i <> manter And Len(ops(i).tipoAux) > 0 Then
            If APS_Ig(ops(i).idPai, idProd) Then
                If ws Is Nothing Then
                    Set ws = APS_Aba(APS_ABA_OPS)
                    If ws Is Nothing Then Exit Function
                    est = APS_Liberar(ws)
                    APS_Ocupado = APS_Ocupado + 1
                End If
                APS_LimparLinha ws, ops(i).linha
                APS_RemoverAuxDaProducao = APS_RemoverAuxDaProducao + 1
            End If
        End If
    Next i
    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Function

'----------------------------------------------------------
' APS_LimparSetupsOrfaos - RECONSTRUCAO COMPATIVEL - NAO CODIGO ORIGINAL,
' mas de alta confianca: reaproveita literalmente o MESMO mecanismo que
' APS_ExcluirOperacao (logo acima, codigo original, intocado) ja usa para
' remover Setup/Limpeza quando a producao e excluida PELA TELA - a mesma
' rotina ja existente APS_LimparLinha (limpa A:J+ID sem excluir a linha
' fisica), o mesmo par tipoAux/idPai, o mesmo APS_Ig para comparar IDs.
'
' Diferenca: APS_ExcluirOperacao ja SABE qual producao sumiu (foi ela quem
' apagou). Esta rotina cobre o caso do comentario original em
' modAPS_Cards.APS_Atualizar ("exclusao manual em 02_Operacoes, por
' exemplo") - por isso varre TODAS as linhas de Setup/Limpeza e verifica,
' para cada uma, se o idPai ainda aponta para uma producao que realmente
' existe em ops(). Se nao existir mais, e orfao: limpa com APS_LimparLinha,
' exatamente como APS_ExcluirOperacao ja faz.
'
' Assinatura CONFIRMADA PELO CODIGO pelo unico ponto de chamada (modAPS_Cards.
' APS_Atualizar): "If APS_LimparSetupsOrfaos(ops, n) Then" (Boolean; n nao e
' reatribuido pela chamadora, que sempre re-le via APS_LerOps depois de True
' - entao n aqui e so leitura).
'
' [ADAPTACAO] Nao remove linhas sem idPai (formato antigo) nem cuja producao ainda esta na planilha.
' Nao mexe em tipoAux, idPai, rotAux, maquina, data, horario, duracao ou
' sequencia de nenhuma linha que NAO seja orfa. Nao chama APS_FecharVaos,
' APS_Encaixar nem cascata.
' [CORRECAO PEN-B-06, PEN-B-20] Limpeza de setups orfaos com backup e rollback.
' Antes de qualquer exclusao, faz backup completo de cada linha considerada orfao.
' Em caso de erro intermediario, restaura todas as linhas backupadas.
' A decisao de considerar uma linha como orfao segue os mesmos criterios de antes
' (idPai inexistente em ops() e na planilha), com a validacao reforçada em
' APS_PaiNaPlanilha (PEN-B-05): a linha encontrada na planilha precisa ser de
' fato uma producao (tipoAux = 0 e sem marcador [[APS_ nas Observacoes)).
Public Function APS_LimparSetupsOrfaos(ByRef ops() As tOperacao, ByVal n As Long) As Boolean
    Dim ws As Worksheet, est As Boolean
    Dim i As Long, j As Long, achouPai As Boolean
    Dim backups() As tBackupLinha, nBk As Long, iBk As Long
    Dim falhaRecuperou As Boolean, msg As String

    On Error GoTo Falha

    If n <= 0 Then Exit Function

    ' [PEN-B-06] Fase 1: coletar backup de cada linha considerada orfao, antes de qualquer alteracao.
    ReDim backups(1 To IIf(n < 1, 1, n))
    nBk = 0
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 And Len(Trim$(ops(i).idPai)) > 0 Then
            achouPai = False
            If Len(Trim$(ops(i).idPai)) > 0 Then
                For j = 1 To n
                    If j <> i Then
                        If Len(ops(j).tipoAux) = 0 Then
                            If APS_Ig(ops(j).id, ops(i).idPai) Then
                                achouPai = True
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If

            ' a producao pode existir na planilha e ter sido apenas ignorada por APS_LerOps
            If Not achouPai Then achouPai = APS_PaiNaPlanilha(ops(i).idPai)

            If Not achouPai Then
                If ws Is Nothing Then
                    Set ws = APS_Aba(APS_ABA_OPS)
                    If ws Is Nothing Then Exit Function
                    est = APS_Liberar(ws)
                    APS_Ocupado = APS_Ocupado + 1
                End If
                nBk = nBk + 1
                If Not APS_BackupLinha(ws, ops(i).linha, backups(nBk)) Then
                    msg = msg & "Linha " & ops(i).linha & ": nao foi possivel fazer backup." & vbCrLf
                End If
            End If
        End If
    Next i

    If Len(msg) > 0 Then
        MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel fazer backup de algumas linhas:" & vbCrLf & msg, vbExclamation, APS_TITULO
    End If

    ' [PEN-B-06] Fase 2: aplica a limpeza.
    For iBk = 1 To nBk
        i = 0
        For i = 1 To n
            If ops(i).linha = backups(iBk).linha Then Exit For
        Next i
        If i > n Then GoTo ProximoBk

        If Len(ops(i).tipoAux) > 0 And Len(Trim$(ops(i).idPai)) > 0 Then
            achouPai = False
            If Len(Trim$(ops(i).idPai)) > 0 Then
                For j = 1 To n
                    If j <> i Then
                        If Len(ops(j).tipoAux) = 0 Then
                            If APS_Ig(ops(j).id, ops(i).idPai) Then
                                achouPai = True
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
            If Not achouPai Then achouPai = APS_PaiNaPlanilha(ops(i).idPai)

            If Not achouPai Then
                APS_LimparLinha ws, ops(i).linha
                APS_LimparSetupsOrfaos = True
            End If
        End If
ProximoBk:
    Next iBk

    ' [PEN-B-20] Sucesso: libera o APS_Ocupado adquirido na Fase 1 antes de sair.
    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
    Exit Function

Falha:
    ' [PEN-B-20] Em caso de erro, restaura todas as linhas backupadas antes de sair.
    On Error Resume Next
    For iBk = 1 To nBk
        APS_RestaurarLinha ws, backups(iBk)
    Next iBk
    MsgBox "Falha durante a limpeza de setups orfaos: " & Err.Description & vbCrLf & _
           "Todas as linhas alteradas foram restauradas ao estado original.", vbCritical, APS_TITULO
    Err.Clear
    falhaRecuperou = True

Fim:
    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Function

'----------------------------------------------------------
' Cabecalho atual - SOMENTE 10 colunas
'----------------------------------------------------------
Private Sub EscreverCabecalho(ByVal ws As Worksheet)

    Dim h As Variant, i As Long, w As Variant

    h = Array( _
        H_COD(), _
        H_PRO(), _
        H_LOT(), _
        H_MAQ(), _
        H_INI(), _
        H_FIM(), _
        H_PLA(), _
        H_OBS(), _
        H_QTD(), _
        H_DUR())

    w = Array(18, 26, 14, 18, 18, 18, 15, 30, 20, 12)

    For i = 0 To 9

        With ws.Cells(1, i + 1)

            .Value = h(i)
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

        ws.Columns(i + 1).ColumnWidth = w(i)

    Next i

    ws.Rows(1).RowHeight = 28

End Sub

'----------------------------------------------------------
' Migracao
'----------------------------------------------------------
Public Function APS_Migrar() As Boolean

    Dim ws As Worksheet, est As Boolean, resp As VbMsgBoxResult
    Dim ops() As tOperacao, n As Long, i As Long

    Set ws = APS_Aba(APS_ABA_OPS)

    If ws Is Nothing Then Exit Function

    If APS_Col(ws, H_COD()) > 0 And _
       APS_Col(ws, H_PRO()) > 0 And _
       APS_Col(ws, H_LOT()) > 0 And _
       APS_Col(ws, H_MAQ()) > 0 And _
       APS_Col(ws, H_INI()) > 0 And _
       APS_Col(ws, H_FIM()) > 0 And _
       APS_Col(ws, H_PLA()) > 0 Then

        APS_Migrar = True
        Exit Function
    End If

    resp = MsgBox( _
        "A aba 02_Operacoes est" & ChrW(225) & _
        " no layout antigo." & vbCrLf & vbCrLf & _
        "Ela ser" & ChrW(225) & " reorganizada para: " & _
        "C" & ChrW(243) & "digo do Produto, Produto, Lote, M" & _
        ChrW(225) & "quina, Data de In" & ChrW(237) & "cio, " & _
        "Data de Fim, Planejamento, Observa" & ChrW(231) & _
        ChrW(245) & "es, Quantidade de Caixas, Dura" & _
        ChrW(231) & ChrW(227) & "o." & vbCrLf & _
        "As colunas antigas 'Atraso' e 'Tempo de Produto' ser" & _
        ChrW(227) & "o descartadas." & vbCrLf & vbCrLf & _
        "Recomendo salvar uma c" & ChrW(243) & _
        "pia do arquivo antes. Continuar?", _
        vbYesNo + vbQuestion, APS_TITULO)

    If resp <> vbYes Then Exit Function

    n = LerLegado(ws, ops)

    est = APS_Liberar(ws)

    APS_Ocupado = APS_Ocupado + 1

    ws.Cells.Clear
    ws.Columns.Hidden = False

    EscreverCabecalho ws

    For i = 1 To n
        ops(i).linha = 0
        APS_GravarOp ops(i)
    Next i

    APS_Ocupado = APS_Ocupado - 1

    APS_Reproteger ws, est

    APS_Migrar = True
End Function

Private Function LerLegado(ByVal ws As Worksheet, _
                           ByRef ops() As tOperacao) As Long

    Dim cOP As Long, cPro As Long, cLot As Long
    Dim cMaq As Long, cDat As Long, cIni As Long, cFim As Long
    Dim cSta As Long, cObs As Long, cMat As Long
    Dim cQtd As Long, cDur As Long

    Dim r As Long, ult As Long, n As Long
    Dim d As Double, h As Double, f As Double, dur As Double
    Dim x As Variant

    Dim hOP As String, hDat As String, hIni As String
    Dim hSta As String, hObs As String, hQtd As String

    hOP = "OP"
    hDat = "Data"
    hIni = "In" & ChrW(237) & "cio"
    hSta = "Status"
    hObs = "Observa" & ChrW(231) & ChrW(227) & "o"
    hQtd = "Quantidade (caixas)"

    cOP = APS_Col(ws, hOP)
    cPro = APS_Col(ws, H_PRO())
    cLot = APS_Col(ws, H_LOT())
    cMaq = APS_Col(ws, H_MAQ())
    cDat = APS_Col(ws, hDat)
    cIni = APS_Col(ws, hIni)
    cFim = APS_Col(ws, "Fim")
    cSta = APS_Col(ws, hSta)
    cObs = APS_Col(ws, hObs)
    cMat = APS_Col(ws, "Material")
    cQtd = APS_Col(ws, hQtd)
    cDur = APS_Col(ws, H_DUR())

    ReDim ops(1 To 1)

    If cPro = 0 Or cMaq = 0 Or cDat = 0 Or cIni = 0 Then Exit Function

    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row

    If ult < 2 Then Exit Function

    ReDim ops(1 To ult - 1)

    For r = 2 To ult

        If Len(APS_Txt(ws.Cells(r, cPro).Value)) > 0 Then

            d = APS_ParaData(ws.Cells(r, cDat).Value)
            h = APS_ParaHora(ws.Cells(r, cIni).Value)

            If d >= 0 And h >= 0 And _
               Len(APS_Txt(ws.Cells(r, cMaq).Value)) > 0 Then

                n = n + 1

                With ops(n)

                    If cOP > 0 Then
                        .id = APS_Txt(ws.Cells(r, cOP).Value)
                    End If

                    If Len(.id) = 0 Then
                        .id = CStr(r)
                    End If

                    .Produto = APS_Txt(ws.Cells(r, cPro).Value)

                    If cMat > 0 Then
                        .codigo = APS_Txt(ws.Cells(r, cMat).Value)
                    End If

                    If Len(.codigo) = 0 Then
                        .codigo = APS_CodigoDoNome(.Produto)
                    End If

                    If Len(APS_NomeDoCodigo(.codigo)) > 0 Then
                        .Produto = APS_NomeDoCodigo(.codigo)
                    End If

                    If cLot > 0 Then
                        .lote = APS_Txt(ws.Cells(r, cLot).Value)
                    End If

                    .maquina = APS_Txt(ws.Cells(r, cMaq).Value)

                    If cSta > 0 Then
                        .status = APS_Txt(ws.Cells(r, cSta).Value)
                    End If

                    If Len(.status) = 0 Then
                        .status = APS_ST_PLANEJADA
                    End If

                    If cObs > 0 Then
                        .obs = APS_Txt(ws.Cells(r, cObs).Value)
                    End If

                    If cQtd > 0 Then

                        x = ws.Cells(r, cQtd).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                .caixas = CDbl(x)
                            End If
                        End If

                    End If

                    dur = 0

                    If cDur > 0 Then

                        x = ws.Cells(r, cDur).Value2

                        If Not IsError(x) Then
                            If IsNumeric(x) And Not IsEmpty(x) Then
                                dur = CDbl(x)
                            End If
                        End If

                    End If

                    If dur <= APS_EPS And cFim > 0 Then

                        f = APS_ParaHora(ws.Cells(r, cFim).Value)

                        If f >= 0 Then
                            dur = f - h

                            If dur < 0 Then dur = dur + 1
                        End If
                    End If

                    dur = APS_ArredMin(dur)

                    .ini = APS_ArredMin(d + h)

                    If dur > APS_EPS Then

                        .dur = dur
                        .Fim = APS_ArredMin(.ini + dur)

                    Else

                        .SemDur = True
                        .Fim = .ini + 1# / 24#

                    End If

                End With
            End If
        End If
    Next r

    LerLegado = n
End Function

'==========================================================
' PREPARAR
'==========================================================
Public Sub APS_Preparar()

    Dim wsO As Worksheet, wsP As Worksheet
    Dim estO As Boolean, estP As Boolean

    On Error GoTo Falha

    Set wsO = APS_Aba(APS_ABA_OPS)
    Set wsP = APS_Aba(APS_ABA_PLAN)

    If wsO Is Nothing Or wsP Is Nothing Then

        MsgBox "As abas 02_Operacoes e 1_Planejamento s" & _
               ChrW(227) & "o necess" & ChrW(225) & "rias.", _
               vbCritical, APS_TITULO

        Exit Sub
    End If

    Application.ScreenUpdating = False

    If Not APS_Migrar() Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    estO = APS_Liberar(wsO)
    estP = APS_Liberar(wsP)

    On Error Resume Next

    wsO.Shapes("BTN_CONFIG_APS").Delete
    wsP.Shapes("TESTE_REGUA_APS").Delete

    On Error GoTo Falha

    wsP.Activate

    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    wsP.Range("B4").Select

    ActiveWindow.FreezePanes = True

    wsP.Range("A1").Select

    APS_Reproteger wsO, estO
    APS_Reproteger wsP, estP

    Application.ScreenUpdating = True

    ' Primeiro atualiza toda a estrutura do planejamento.
    ' Depois recria os botoes para que + ADICIONAR DIAS
    ' seja posicionado com base no layout final.
    APS_Atualizar True
    APS_CriarBotoes
    APS_IrParaHoje True

    MsgBox "Estrutura preparada.", vbInformation, APS_TITULO

    Exit Sub

Falha:

    On Error Resume Next

    APS_Reproteger wsO, estO
    APS_Reproteger wsP, estP

    Application.ScreenUpdating = True

    MsgBox "Erro ao preparar a estrutura: " & Err.Description, _
           vbCritical, APS_TITULO
End Sub

' Chamado pelo Workbook_Open
Public Sub APS_AoAbrir()

    Dim falhas As String

    APS_Ocupado = 0

    ' Cada etapa e isolada: uma falha nao impede as demais e e REPORTADA no fim (nada de erro engolido).
    On Error Resume Next
    APS_PrepararMes
    If Err.Number <> 0 Then falhas = falhas & "- Preparar mes: " & Err.Description & vbCrLf: Err.Clear

    ' aba SAP (so cria se ainda nao existir)
    APS_GarantirAbaSAP
    If Err.Number <> 0 Then falhas = falhas & "- Aba SAP: " & Err.Description & vbCrLf: Err.Clear

    ' cabecalho azul, dados normais (limpa azul copiado para linhas de dados)
    APS_CorrigirFormatoDados
    If Err.Number <> 0 Then falhas = falhas & "- Formato dos dados: " & Err.Description & vbCrLf: Err.Clear

    ' status automatico + Setup/Limpeza + cards
    APS_Atualizar True
    If Err.Number <> 0 Then falhas = falhas & "- Atualizar: " & Err.Description & vbCrLf: Err.Clear
    APS_IrParaHoje True
    If Err.Number <> 0 Then falhas = falhas & "- Ir para hoje: " & Err.Description & vbCrLf: Err.Clear

    ' botao CADASTROS E PARAMETROS na tela de planejamento (so cria se faltar; abaixo de HOJE/CALENDARIO)
    APS_GarantirBotaoParametros
    If Err.Number <> 0 Then falhas = falhas & "- Botao Parametros: " & Err.Description & vbCrLf: Err.Clear

    ' aba Visao Geral Mensal (somente consulta; abre no mes e ano atuais)
    APS_VM_Iniciar
    If Err.Number <> 0 Then falhas = falhas & "- Visao Geral Mensal: " & Err.Description & vbCrLf: Err.Clear

    ' relogio: confere o status a cada minuto
    APS_AgendarTick

    ' FASE 6: sincronizacao inicial (so em modo conectado; no-op em modo local)
    Central_Sincronizar True
    If Err.Number <> 0 Then falhas = falhas & "- Sincronizacao central: " & Err.Description & vbCrLf: Err.Clear
    On Error GoTo 0

    If Len(falhas) > 0 Then MsgBox "O APS abriu, mas algumas etapas falharam:" & vbCrLf & falhas, vbExclamation, APS_TITULO

End Sub


' Chamada por Workbook_SheetActivate: confere se algo mudou (ex.: linha excluida) e redesenha.

Public Sub APS_SincronizarSeMudou()
    If APS_Ocupado > 0 Then Exit Sub
    On Error Resume Next
    If APS_PrecisaAtualizar() Then APS_Atualizar True
    APS_GarantirTick
End Sub

' Chamado por Workbook_SheetChange (02_Operacoes)
Public Sub APS_OpsAlterada(ByVal Target As Range)
    Dim ws As Worksheet, rel As Range, cel As Range, est As Boolean, trav As Boolean
    Dim cIni As Long, cFim As Long, cDur As Long, cMaq As Long, cPla As Long, cSet As Long, cAtr As Long
    Dim bitsR() As Long, rmin As Long, rmax As Long, ar As Range, i As Long, r As Long, b As Long
    Dim msg As String, achou As Boolean

    If APS_Ocupado > 0 Then Exit Sub
    On Error GoTo Falha

    Set ws = Target.Worksheet
    Set rel = Intersect(Target, ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, APS_NColDados(ws))))
    If rel Is Nothing Then Exit Sub

    ' FASE 6: edicao manual direta na planilha tambem fica pendente de envio (no-op em modo local)
    If Central_Ativo() Then
        For Each cel In rel.Cells
            Central_MarcarLinha cel.Row, "UPSERT"
        Next cel
    End If

    APS_Ocupado = APS_Ocupado + 1
    trav = True
    est = APS_Liberar(ws)

    cIni = APS_Col(ws, H_INI()): cFim = APS_Col(ws, H_FIM()): cDur = APS_Col(ws, H_DUR())
    cMaq = APS_Col(ws, H_MAQ()): cPla = APS_Col(ws, H_PLA()): cSet = APS_Col(ws, H_SET()): cAtr = APS_Col(ws, H_ATR())

    ' Sem limite de linhas/celulas (colar um bloco grande nao descarta edicoes): bits por numero de linha.
    rmin = ws.Rows.Count: rmax = 0
    For Each ar In rel.Areas
        If ar.Row < rmin Then rmin = ar.Row
        If ar.Row + ar.Rows.Count - 1 > rmax Then rmax = ar.Row + ar.Rows.Count - 1
    Next ar
    ReDim bitsR(rmin To rmax)

    For Each cel In rel.Cells
        r = cel.Row
        b = 0
        If cel.Column = cIni Then b = 1
        If cel.Column = cFim Then b = 2
        If cel.Column = cDur Then b = 4
        If cSet > 0 Then
            If cel.Column = cSet Then
                b = 4 Or 64      ' mudou o Setup: o bloco muda de duracao (64 = o INICIO DA PREPARACAO fica; a producao anda)
                ' numero >= 1 digitado = minutos (ex.: 90): vira tempo hh:mm
                If VarType(cel.Value2) = vbDouble Then
                    If cel.Value2 >= 1# Then
                        cel.NumberFormat = "[h]:mm"
                        cel.Value2 = APS_ArredMin(cel.Value2 / 1440#)
                    End If
                End If
            End If
        End If
        If cAtr > 0 Then
            If cel.Column = cAtr Then
                b = 4      ' mudou o Atraso: muda o tempo total programado
                If VarType(cel.Value2) = vbDouble Then
                    If cel.Value2 >= 1# Then
                        cel.NumberFormat = "[h]:mm"
                        cel.Value2 = APS_ArredMin(cel.Value2 / 1440#)
                    End If
                End If
            End If
        End If
        If cel.Column = cMaq Then b = 8
        If cel.Column = cPla Then b = 16
        bitsR(r) = bitsR(r) Or b Or 32      ' 32 = linha tocada (formata/reencaixa conforme os outros bits)
    Next cel

    For r = rmin To rmax
        If bitsR(r) <> 0 Then
            If Len(APS_Txt(ws.Cells(r, 1).Value)) > 0 Or Len(APS_Txt(ws.Cells(r, 2).Value)) > 0 Then
                If ws.Cells(r, 1).Interior.Color = APS_COR_CAB Then APS_FormatarLinhaDados ws, r
                If (bitsR(r) And 15) <> 0 Then APS_ReencaixarLinha r, bitsR(r)
                If (bitsR(r) And 16) <> 0 Then MarcarStatusLinha r
            End If
        End If
    Next r

    APS_Ocupado = APS_Ocupado - 1
    trav = False

    ' recompoe status, Setup/Limpeza e cards (tambem limpa card/ID de linha excluida)
    APS_Atualizar True
    APS_Reproteger ws, est
    Exit Sub

Falha:
    msg = Err.Description
    If trav Then APS_Ocupado = APS_Ocupado - 1
    On Error Resume Next
    If est Then APS_Reproteger ws, est
    MsgBox "Erro ao aplicar a altera" & ChrW(231) & ChrW(227) & "o em 02_Operacoes: " & msg, vbExclamation, APS_TITULO
End Sub

Private Sub MarcarStatusLinha(ByVal linha As Long)
    Dim ops() As tOperacao, n As Long, i As Long
    n = APS_LerOps(ops)
    For i = 1 To n
        If ops(i).linha = linha Then
            APS_MapaMarcar linha, ops(i).ini, ops(i).Fim
            Exit For
        End If
    Next i
End Sub

' Normaliza o que foi digitado em A1 para o dia 1 do mes
Public Sub APS_MesAlterado(ByVal Target As Range)
    Dim ws As Worksheet, v As Variant, d As Date, alvo As Double, est As Boolean

    If APS_Ocupado > 0 Then Exit Sub
    Set ws = Target.Worksheet
    If Intersect(Target, ws.Range("A1")) Is Nothing Then Exit Sub
    On Error GoTo Sai
    v = ws.Range("A1").Value2
    If IsError(v) Then Exit Sub
    If VarType(v) <> vbDouble Then Exit Sub
    If v < 36526 Then Exit Sub
    d = CDate(Int(v))
    alvo = CDbl(DateSerial(Year(d), Month(d), 1))
    If Abs(v - alvo) > 0.000001 Then
        APS_Ocupado = APS_Ocupado + 1
        est = APS_Liberar(ws)
        ws.Range("A1").Value2 = alvo
        APS_Ocupado = APS_Ocupado - 1
    End If
Sai:
    If Err.Number <> 0 Then
        Err.Clear
        If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    End If
    On Error Resume Next
    If est Then APS_Reproteger ws, est
End Sub


'==========================================================
' Colunas de dados / coluna OP
'==========================================================
' Ultima coluna de DADOS da operacao = tudo antes da coluna ID tecnica. Substitui o "10" fixo, para que a
' coluna OP (e qualquer outra inserida antes do ID) entre na limpeza, na assinatura, no formato e nos eventos.
Public Function APS_NColDados(ByVal ws As Worksheet) As Long
    Dim c As Long

    c = APS_Col(ws, H_ID())
    If c > 1 Then
        APS_NColDados = c - 1
    Else
        APS_NColDados = 10
    End If
End Function

' Garante a coluna OP logo depois de Produto. Idempotente: so insere se ainda nao existir. A insercao e feita
' pelo proprio Excel (as formulas de outras abas acompanham). Nenhuma outra coluna e reordenada.
Public Sub APS_GarantirColunaOP()
    If Central_EstruturaTravada() Then Exit Sub     ' FASE 6
    Dim ws As Worksheet, cPro As Long, est As Boolean, ult As Long, trav As Boolean

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    cPro = APS_Col(ws, H_PRO())
    If cPro = 0 Then Exit Sub                       ' layout antigo: a migracao cuida
    If APS_Col(ws, H_INI()) = 0 Then Exit Sub
    If APS_Col(ws, H_ORD()) > 0 Then Exit Sub       ' ja existe

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    trav = True

    ws.Columns(cPro + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, cPro).Copy
    ws.Cells(1, cPro + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(1, cPro + 1).Value = H_ORD()
    ws.Columns(cPro + 1).ColumnWidth = 14
    ult = ws.Cells(ws.Rows.Count, cPro).End(xlUp).Row
    If ult >= 2 Then ws.Range(ws.Cells(2, cPro + 1), ws.Cells(ult, cPro + 1)).NumberFormat = "@"

Sai:
    Application.CutCopyMode = False
    If trav Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Sub

'==========================================================
' Coluna Setup (tempo de Setup/Limpeza na propria linha da operacao)
'==========================================================
' Garante a coluna "Setup" logo depois de Duracao. Idempotente: so insere se ainda nao existir. A insercao e
' feita pelo proprio Excel (formulas de outras abas acompanham). Nenhuma outra coluna e reordenada.
Public Sub APS_GarantirColunaSetup()
    If Central_EstruturaTravada() Then Exit Sub     ' FASE 6
    Dim ws As Worksheet, cDur As Long, est As Boolean, ult As Long, trav As Boolean

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    If APS_Col(ws, H_INI()) = 0 Then Exit Sub        ' layout antigo: a migracao cuida
    cDur = APS_Col(ws, H_DUR())
    If cDur = 0 Then Exit Sub
    If APS_Col(ws, H_SET()) > 0 Then Exit Sub        ' ja existe

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    trav = True

    ws.Columns(cDur + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, cDur).Copy
    ws.Cells(1, cDur + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(1, cDur + 1).Value = H_SET()
    ws.Columns(cDur + 1).Hidden = False
    ws.Columns(cDur + 1).ColumnWidth = 10
    ult = ws.Cells(ws.Rows.Count, APS_Col(ws, H_PRO())).End(xlUp).Row
    If ult >= 2 Then ws.Range(ws.Cells(2, cDur + 1), ws.Cells(ult, cDur + 1)).NumberFormat = "[h]:mm"

Sai:
    Application.CutCopyMode = False
    If trav Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Sub

' Garante a coluna "Atraso" logo depois de Setup. Idempotente; insercao feita pelo Excel; nenhuma outra coluna e reordenada.
Public Sub APS_GarantirColunaAtraso()
    If Central_EstruturaTravada() Then Exit Sub     ' FASE 6
    Dim ws As Worksheet, cDur As Long, est As Boolean, ult As Long, trav As Boolean

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    If APS_Col(ws, H_INI()) = 0 Then Exit Sub        ' layout antigo: a migracao cuida
    cDur = APS_Col(ws, H_SET())   ' (cDur = coluna Setup: o Atraso entra logo depois)
    If cDur = 0 Then Exit Sub        ' sem coluna Setup ainda: nada a fazer agora
    If APS_Col(ws, H_ATR()) > 0 Then Exit Sub        ' ja existe

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    trav = True

    ws.Columns(cDur + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, cDur).Copy
    ws.Cells(1, cDur + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(1, cDur + 1).Value = H_ATR()
    ws.Columns(cDur + 1).Hidden = False
    ws.Columns(cDur + 1).ColumnWidth = 10
    ult = ws.Cells(ws.Rows.Count, APS_Col(ws, H_PRO())).End(xlUp).Row
    If ult >= 2 Then ws.Range(ws.Cells(2, cDur + 1), ws.Cells(ult, cDur + 1)).NumberFormat = "[h]:mm"

Sai:
    Application.CutCopyMode = False
    If trav Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Sub

' Recolhe as linhas de Setup/Limpeza do formato com linha propria (marcador [[APS_..._DE:<id>]]): o tempo vai
' para a coluna Setup da producao-pai e a linha de Setup e limpa (mesma limpeza de APS_ExcluirOperacao).
'==========================================================
' [CORRECAO PEN-B-22] Estrutura usada para backup e rollback de
' linhas afetadas pela migracao de setups. Armazena os valores
' e formatos de cada celula da linha, permitindo restauracao
' completa em caso de falha.
'==========================================================
' Faz backup completo de uma linha (valores e formatos de todas as colunas ateh a ultima com cabecalho).
Private Function APS_BackupLinha(ByVal ws As Worksheet, ByVal linha As Long, ByRef b As tBackupLinha) As Boolean
    Dim c As Long, ult As Long

    On Error GoTo Falha

    If linha < 2 Then Exit Function

    b.linha = linha
    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ult < 1 Then ult = APS_NColDados(ws)
    b.nCol = ult
    ReDim b.valores(1 To ult)
    ReDim b.formatos(1 To ult)

    For c = 1 To ult
        b.valores(c) = ws.Cells(linha, c).Value2
        b.formatos(c) = ws.Cells(linha, c).NumberFormat
    Next c

    APS_BackupLinha = True
    Exit Function

Falha:
    APS_BackupLinha = False
End Function

' Restaura uma linha a partir do backup salvo. Sempre tenta, mesmo que alguma celula falhe.
Private Sub APS_RestaurarLinha(ByVal ws As Worksheet, ByRef b As tBackupLinha)
    Dim c As Long

    On Error Resume Next

    If b.linha < 2 Then Exit Sub
    For c = 1 To b.nCol
        ws.Cells(b.linha, c).Value2 = b.valores(c)
        ws.Cells(b.linha, c).NumberFormat = b.formatos(c)
    Next c
End Sub

' Corrige a origem da duplicidade; nao esconde nada. Sem producao-pai encontrada: a linha nao e tocada.
' [CORRECAO PEN-B-22] Antes de qualquer alteracao, faz backup completo de cada linha de Setup
' que sera migrada. Em caso de erro, restaura todas as linhas backupadas. A migracao e idempotente:
' se a coluna Setup da producao ja tiver valor, o tempo nao e sobrescrito; se a linha de Setup
' ja estiver vazia, APS_LerOps a ignora na proxima execucao.
Public Sub APS_MigrarSetupsParaColuna()
    If Central_EstruturaTravada() Then Exit Sub     ' FASE 6
    Dim ws As Worksheet, ops() As tOperacao, n As Long, i As Long, j As Long, pai As Long
    Dim cSet As Long, cObs As Long, est As Boolean, trav As Boolean, tem As Boolean, ult As Long, v As Variant
    Dim backups() As tBackupLinha, nBk As Long, iBk As Long, falhaRecuperou As Boolean

    On Error GoTo Sai

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    cSet = APS_Col(ws, H_SET())
    cObs = APS_Col(ws, H_OBS())
    If cSet = 0 Or cObs = 0 Then Exit Sub

    ' verificacao barata: so le as operacoes se alguma Observacao tiver o marcador de vinculo
    ult = ws.Cells(ws.Rows.Count, cObs).End(xlUp).Row
    If ult < 2 Then Exit Sub
    v = ws.Range(ws.Cells(2, cObs), ws.Cells(ult, cObs)).Value2
    If Not IsArray(v) Then
        If Not IsError(v) Then tem = (InStr(1, CStr(v), "[[APS_", vbTextCompare) > 0)
    Else
        For i = 1 To UBound(v, 1)
            If Not IsError(v(i, 1)) Then
                If InStr(1, CStr(v(i, 1)), "[[APS_", vbTextCompare) > 0 Then tem = True: Exit For
            End If
        Next i
    End If
    If Not tem Then Exit Sub

    n = APS_LerOps(ops)
    If n <= 0 Then Exit Sub

    ' [PEN-B-22] Fase 1: coletar backup de cada linha de Setup que sera migrada, antes de qualquer alteracao.
    ReDim backups(1 To IIf(n < 1, 1, n))
    nBk = 0
    For i = 1 To n
        If ops(i).linha > 0 And Len(ops(i).tipoAux) > 0 And Len(Trim$(ops(i).idPai)) > 0 Then
            pai = 0
            For j = 1 To n
                If ops(j).linha > 0 And Len(ops(j).tipoAux) = 0 Then
                    If APS_Ig(ops(j).id, ops(i).idPai) Then pai = j: Exit For
                End If
            Next j
            If pai > 0 Then
                nBk = nBk + 1
                If Not APS_BackupLinha(ws, ops(i).linha, backups(nBk)) Then
                    ' nao foi possivel fazer backup: nao pode migrar com seguranca
                    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel fazer backup da linha " & ops(i).linha & _
                           " antes da migra" & ChrW(231) & ChrW(227) & "o. A migra" & ChrW(231) & ChrW(227) & "o foi cancelada para evitar perda de dados.", _
                           vbExclamation, APS_TITULO
                    Exit Sub
                End If
            End If
        End If
    Next i

    ' [PEN-B-22] Fase 2: aplica a migracao. Se qualquer etapa falhar, restaura todas as linhas backupadas.
    For iBk = 1 To nBk
        i = 0
        ' localiza o indice correspondente na array ops
        For i = 1 To n
            If ops(i).linha = backups(iBk).linha Then Exit For
        Next i
        If i > n Then GoTo ProximoBk

        If ops(i).linha > 0 And Len(ops(i).tipoAux) > 0 And Len(Trim$(ops(i).idPai)) > 0 Then
            pai = 0
            For j = 1 To n
                If ops(j).linha > 0 And Len(ops(j).tipoAux) = 0 Then
                    If APS_Ig(ops(j).id, ops(i).idPai) Then pai = j: Exit For
                End If
            Next j
            If pai > 0 Then
                If Not trav Then
                    est = APS_Liberar(ws)
                    APS_Ocupado = APS_Ocupado + 1
                    trav = True
                End If
                ' o tempo so entra se a producao ainda nao tiver Setup na coluna
                If APS_SetupDoValor(ws.Cells(ops(pai).linha, cSet).Value2) <= APS_EPS Then
                    APS_GravarSetupDaOp ws, ops(pai).linha, ops(i).dur
                End If
                APS_LimparLinha ws, ops(i).linha
            End If
        End If
ProximoBk:
    Next iBk

Sai:
    On Error Resume Next
    ' [PEN-B-22] Em caso de erro, restaura todas as linhas backupadas antes de sair.
    If Err.Number <> 0 Then
        For iBk = 1 To nBk
            APS_RestaurarLinha ws, backups(iBk)
        Next iBk
        MsgBox "Falha durante a migra" & ChrW(231) & ChrW(227) & "o de setups: " & Err.Description & vbCrLf & _
               "Todas as linhas alteradas foram restauradas ao estado original.", vbCritical, APS_TITULO
        Err.Clear
        falhaRecuperou = True
    End If
    If trav And Not falhaRecuperou Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    ElseIf trav Then
        ' em caso de falha, restaura tambem as travas e protecoes
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Sub

'==========================================================
' Dashboard (00_Dashboard). Corrige os indicadores da auditoria sem alterar o layout:
'  - status contados na coluna Planejamento (e nao em Observacoes)
'  - Total de OPs e demais contagens: so PRODUCAO (Setup/Limpeza reconhecidos pelo marcador [[APS_...]]
'    nas Observacoes ou pelo Codigo SETUP/LIMPEZA); sem limite de 100 linhas (colunas inteiras)
'  - Horas Totais = soma das DURACOES das producoes (em horas)
'  - Ocupacao Media: ver APS_DashOcupacaoMedia
'==========================================================
Private Function APS_RngCol(ByVal ws As Worksheet, ByVal c As Long) As String
    Dim l As String

    l = Split(ws.Cells(1, c).Address(True, False), "$")(0)
    APS_RngCol = "'" & ws.Name & "'!$" & l & ":$" & l
End Function

Public Sub APS_AtualizarDashboard()
    Dim wsD As Worksheet, wsO As Worksheet
    Dim cIni As Long, cCod As Long, cObs As Long, cPla As Long, cDur As Long
    Dim crit As String, q As String, f As String

    On Error GoTo Sai

    Set wsO = APS_Aba(APS_ABA_OPS)
    If wsO Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets("00_Dashboard")
    On Error GoTo Sai
    If wsD Is Nothing Then Exit Sub

    cIni = APS_Col(wsO, H_INI()): cCod = APS_Col(wsO, H_COD()): cObs = APS_Col(wsO, H_OBS())
    cPla = APS_Col(wsO, H_PLA()): cDur = APS_Col(wsO, H_DUR())
    If cIni = 0 Or cCod = 0 Or cObs = 0 Or cPla = 0 Or cDur = 0 Then Exit Sub

    q = Chr$(34)
    ' linha real de operacao (Inicio numerico) e que NAO e Setup/Limpeza
    crit = APS_RngCol(wsO, cIni) & "," & q & ">0" & q & "," & _
           APS_RngCol(wsO, cObs) & "," & q & "<>*[[APS_*" & q & "," & _
           APS_RngCol(wsO, cCod) & "," & q & "<>" & APS_COD_SETUP & q & "," & _
           APS_RngCol(wsO, cCod) & "," & q & "<>" & APS_COD_LIMPEZA & q

    APS_DashFormula wsD, "A4", "=COUNTIFS(" & crit & ")"
    APS_DashFormula wsD, "C4", "=COUNTIFS(" & crit & "," & APS_RngCol(wsO, cPla) & "," & q & APS_ST_PLANEJADA & q & ")"
    APS_DashFormula wsD, "E4", "=COUNTIFS(" & crit & "," & APS_RngCol(wsO, cPla) & "," & q & APS_ST_ANDAMENTO & q & ")"
    APS_DashFormula wsD, "G4", "=COUNTIFS(" & crit & "," & APS_RngCol(wsO, cPla) & "," & q & APS_ST_CONCLUIDA() & q & ")"
    APS_DashFormula wsD, "I4", "=COUNTIFS(" & crit & "," & APS_RngCol(wsO, cPla) & "," & q & APS_ST_ATRASADA & q & ")"
    APS_DashFormula wsD, "K4", "=SUMIFS(" & APS_RngCol(wsO, cDur) & "," & crit & ")*24"
    APS_DashFormula wsD, "M4", "=APS_DashOcupacaoMedia()"
    wsD.Range("K4").NumberFormat = "0.0"
    wsD.Range("M4").NumberFormat = "0%"

Sai:
End Sub

Private Sub APS_DashFormula(ByVal ws As Worksheet, ByVal addr As String, ByVal f As String)
    If ws.Range(addr).Formula <> f Then ws.Range(addr).Formula = f
End Sub

' Ocupacao media das maquinas: para cada maquina com operacoes ativas,
'   horas de trabalho ocupadas (producao + Setup/Limpeza, pelo calendario) /
'   horas de trabalho disponiveis entre o primeiro inicio e o ultimo fim da maquina (pelo calendario);
' media simples entre as maquinas (cada uma limitada a 100%). DEFINICAO ADOTADA: a planilha nao trazia
' uma regra de ocupacao (a formula anterior era a media das duracoes).
Public Function APS_DashOcupacaoMedia() As Variant
    Dim ops() As tOperacao, n As Long, i As Long, j As Long, k As Long
    Dim maq() As String, nm As Long, achou As Boolean
    Dim ocup As Double, mi As Double, mf As Double, disp As Double, soma As Double, cont As Long

    Application.Volatile
    On Error GoTo Falha

    APS_CalRecarregar
    n = APS_LerOps(ops)
    If n <= 0 Then
        APS_DashOcupacaoMedia = 0
        Exit Function
    End If

    ReDim maq(1 To n)
    For i = 1 To n
        If APS_Ativa(ops(i)) And Len(Trim$(ops(i).maquina)) > 0 Then
            achou = False
            For j = 1 To nm
                If APS_Ig(maq(j), ops(i).maquina) Then achou = True: Exit For
            Next j
            If Not achou Then
                nm = nm + 1
                maq(nm) = ops(i).maquina
            End If
        End If
    Next i

    For k = 1 To nm
        ocup = 0: mi = 0: mf = 0
        For i = 1 To n
            If APS_Ativa(ops(i)) Then
                If APS_Ig(ops(i).maquina, maq(k)) Then
                    ocup = ocup + APS_TrabalhoEntre(ops(i).ini, ops(i).Fim)
                    If mi = 0 Or ops(i).ini < mi Then mi = ops(i).ini
                    If ops(i).Fim > mf Then mf = ops(i).Fim
                End If
            End If
        Next i
        disp = APS_TrabalhoEntre(mi, mf)
        If disp > APS_EPS Then
            If ocup > disp Then ocup = disp
            soma = soma + ocup / disp
            cont = cont + 1
        End If
    Next k

    If cont > 0 Then APS_DashOcupacaoMedia = soma / cont Else APS_DashOcupacaoMedia = 0
    Exit Function

Falha:
    APS_DashOcupacaoMedia = CVErr(xlErrValue)
End Function

'==========================================================
' APS_MigrarSetupsAntigos - MANUAL (Alt+F8). Migra Setup/Limpeza do formato ANTIGO
' (Codigo = SETUP, Produto = rotulo, Lote = "SETUP:<id>") para o padrao novo, e limpa o status gravado.
' So altera, nessas linhas: Codigo, Produto, Lote, Observacoes e Planejamento (status). NAO altera ID,
' vinculo (idPai), horarios, duracao nem maquina. Linhas sem producao-pai encontrada sao ignoradas.
'==========================================================
Public Sub APS_MigrarSetupsAntigos()
    Dim ws As Worksheet, ops() As tOperacao, n As Long, i As Long, j As Long, pai As Long
    Dim cCod As Long, cPro As Long, cLot As Long, cObs As Long, cPla As Long
    Dim antigo As Boolean, comStatus As Boolean, tipo As String, rot As String
    Dim nCand As Long, nSemPai As Long, nFeitas As Long, resp As VbMsgBoxResult, est As Boolean
    Dim marc() As Boolean

    On Error GoTo Falha

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    n = APS_LerOps(ops)
    If n <= 0 Then Exit Sub
    cCod = APS_Col(ws, H_COD()): cPro = APS_Col(ws, H_PRO()): cLot = APS_Col(ws, H_LOT())
    cObs = APS_Col(ws, H_OBS()): cPla = APS_Col(ws, H_PLA())
    If cCod = 0 Or cPro = 0 Or cLot = 0 Or cObs = 0 Or cPla = 0 Then Exit Sub

    ReDim marc(1 To n)
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 And Len(Trim$(ops(i).idPai)) > 0 And ops(i).linha > 0 Then
            antigo = (UCase$(Left$(Trim$(ops(i).lote), 6)) = "SETUP:") Or (UCase$(Left$(Trim$(ops(i).lote), 9)) = "LIMPEZA:")
            comStatus = (Len(Trim$(CStr(ws.Cells(ops(i).linha, cPla).Value))) > 0)
            If antigo Or comStatus Then
                pai = 0
                For j = 1 To n
                    If Len(ops(j).tipoAux) = 0 Then
                        If APS_Ig(ops(j).id, ops(i).idPai) Then pai = j: Exit For
                    End If
                Next j
                If pai = 0 Then
                    nSemPai = nSemPai + 1
                Else
                    marc(i) = True
                    nCand = nCand + 1
                End If
            End If
        End If
    Next i

    If nCand = 0 Then
        MsgBox "Nenhum Setup/Limpeza no formato antigo ou com status gravado." & _
               IIf(nSemPai > 0, vbCrLf & nSemPai & " linha(s) sem produ" & ChrW(231) & ChrW(227) & "o-pai foram ignoradas.", ""), vbInformation, APS_TITULO
        Exit Sub
    End If

    resp = MsgBox(nCand & " linha(s) de Setup/Limpeza ser" & ChrW(227) & "o migradas para o padr" & ChrW(227) & "o novo." & vbCrLf & vbCrLf & _
                  "Ser" & ChrW(227) & "o alterados SOMENTE: C" & ChrW(243) & "digo, Produto, Lote, Observa" & ChrW(231) & ChrW(245) & "es e Planejamento (status)." & vbCrLf & _
                  "ID, v" & ChrW(237) & "nculo, hor" & ChrW(225) & "rios, dura" & ChrW(231) & ChrW(227) & "o e m" & ChrW(225) & "quina N" & ChrW(195) & "O mudam." & vbCrLf & vbCrLf & _
                  "Salve uma c" & ChrW(243) & "pia do arquivo antes. Continuar?", vbYesNo + vbQuestion + vbDefaultButton2, APS_TITULO)
    If resp <> vbYes Then Exit Sub

    est = APS_Liberar(ws)
    APS_Ocupado = APS_Ocupado + 1
    For i = 1 To n
        If marc(i) Then
            pai = 0
            For j = 1 To n
                If Len(ops(j).tipoAux) = 0 Then
                    If APS_Ig(ops(j).id, ops(i).idPai) Then pai = j: Exit For
                End If
            Next j
            rot = ops(i).rotAux
            If Len(rot) = 0 Then rot = ops(i).Produto
            If InStr(1, rot, "Limpeza", vbTextCompare) > 0 Then tipo = APS_TIPO_LIMPEZA Else tipo = APS_TIPO_SETUP

            If (UCase$(Left$(Trim$(ops(i).lote), 6)) = "SETUP:") Or (UCase$(Left$(Trim$(ops(i).lote), 9)) = "LIMPEZA:") Then
                ws.Cells(ops(i).linha, cCod).NumberFormat = "@"
                If tipo = APS_TIPO_LIMPEZA Then ws.Cells(ops(i).linha, cCod).Value = APS_COD_LIMPEZA Else ws.Cells(ops(i).linha, cCod).Value = APS_COD_SETUP
                ws.Cells(ops(i).linha, cPro).Value = ops(pai).Produto
                ws.Cells(ops(i).linha, cLot).NumberFormat = "@"
                ws.Cells(ops(i).linha, cLot).Value = ops(pai).lote
                ws.Cells(ops(i).linha, cObs).Value = APS_MarcadorVinculo(tipo, ops(i).idPai) & " " & rot & " - " & ops(i).maquina
            End If
            ws.Cells(ops(i).linha, cPla).ClearContents
            nFeitas = nFeitas + 1
        End If
    Next i
    APS_Ocupado = APS_Ocupado - 1
    APS_Reproteger ws, est

    APS_Atualizar True
    MsgBox nFeitas & " linha(s) migrada(s)." & IIf(nSemPai > 0, vbCrLf & nSemPai & " ignorada(s) por n" & ChrW(227) & "o terem produ" & ChrW(231) & ChrW(227) & "o-pai.", ""), vbInformation, APS_TITULO
    Exit Sub

Falha:
    On Error Resume Next
    If APS_Ocupado > 0 Then APS_Ocupado = APS_Ocupado - 1
    APS_Reproteger ws, est
    MsgBox "Erro na migra" & ChrW(231) & ChrW(227) & "o: " & Err.Description, vbExclamation, APS_TITULO
End Sub

'==========================================================
' APS_AuditarVinculos - DIAGNOSTICO. Nao altera nenhuma operacao nem card.
' Gera (e recria a cada execucao) a aba APS_AUDITORIA: uma linha por linha de 02_Operacoes com os
' achados do vinculo Producao <-> Setup/Limpeza, e os cards do planejamento que nao tem operacao.
' Execucao: Alt+F8 > APS_AuditarVinculos. A aba pode ser apagada depois.
'==========================================================
Public Sub APS_AuditarVinculos()
    Dim ops() As tOperacao, n As Long, i As Long, j As Long, p As Long, r As Long
    Dim wsO As Worksheet, wsP As Worksheet, wsA As Worksheet, shp As Shape
    Dim cPla As Long, achado As String, nomes As String, ids As String
    Dim nProd As Long, nAux As Long, nProv As Long, nProb As Long, nSob As Long, qtd As Long
    Dim dif As Double, achou As Boolean, cab As Variant, c As Long, idCard As String
    Dim info As String, idsVistos As String, cPro As Long, cOrd As Long

    On Error GoTo Falha

    Set wsO = APS_Aba(APS_ABA_OPS)
    Set wsP = APS_Aba(APS_ABA_PLAN)
    If wsO Is Nothing Or wsP Is Nothing Then Exit Sub

    APS_CalRecarregar
    n = APS_LerOps(ops)
    If n <= 0 Then
        MsgBox "Nenhuma opera" & ChrW(231) & ChrW(227) & "o lida em 02_Operacoes.", vbInformation, APS_TITULO
        Exit Sub
    End If
    cPla = APS_Col(wsO, H_PLA())

    ' cards existentes no planejamento (nome = OPCARD_ + ID)
    For Each shp In wsP.Shapes
        If Left$(shp.Name, 7) = "OPCARD_" Then nomes = nomes & "|" & shp.Name & "|"
    Next shp
    For i = 1 To n
        ids = ids & "|OPCARD_" & ops(i).id & "|"
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("APS_AUDITORIA").Delete
    On Error GoTo Falha
    Application.DisplayAlerts = True
    Set wsA = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsA.Name = "APS_AUDITORIA"

    cab = Array("Tipo", "Linha", "ID", "Produto", "Lote", "M" & ChrW(225) & "quina", "In" & ChrW(237) & "cio", "Fim", _
                "Dura" & ChrW(231) & ChrW(227) & "o (min)", "Status lido", "idPai", "Achado", "OP")
    For c = 0 To UBound(cab)
        wsA.Cells(4, c + 1).Value = cab(c)
    Next c
    wsA.Range(wsA.Cells(4, 1), wsA.Cells(4, UBound(cab) + 1)).Font.Bold = True

    r = 4
    For i = 1 To n
        achado = "": info = ""

        ' ID tecnico duplicado (independente de OP, lote e linha)
        If InStr(idsVistos, "|" & ops(i).id & "|") > 0 Then achado = achado & "ID DUPLICADO; "
        idsVistos = idsVistos & "|" & ops(i).id & "|"

        If Len(ops(i).tipoAux) = 0 Then
            nProd = nProd + 1
            qtd = 0: p = 0
            For j = 1 To n
                If Len(ops(j).tipoAux) > 0 Then
                    If APS_Ig(ops(j).idPai, ops(i).id) Then
                        qtd = qtd + 1
                        If p = 0 Then p = j
                    End If
                End If
            Next j
            If qtd > 1 Then achado = achado & "MAIS DE UM Setup/Limpeza vinculado (" & qtd & "); "
            ' OP (informativo: a mesma OP pode aparecer em mais de uma etapa/maquina)
            If Len(Trim$(ops(i).ordem)) = 0 Then
                info = info & "sem OP informada; "
            Else
                For j = i + 1 To n
                    If Len(ops(j).tipoAux) = 0 Then
                        If StrComp(Trim$(ops(j).ordem), Trim$(ops(i).ordem), vbTextCompare) = 0 Then info = info & "OP repetida na linha " & ops(j).linha & "; "
                    End If
                Next j
            End If
            If APS_EhMediseal(ops(i).maquina) Then
                If ops(i).oee <= 0 Or ops(i).velBase <= 0 Then info = info & "OEE/velocidade-base n" & ChrW(227) & "o gravados nesta opera" & ChrW(231) & ChrW(227) & "o; "
            End If
            If p > 0 Then
                dif = (ops(p).Fim - ops(i).ini) * 1440#
                If Abs(dif) > 0.5 Then achado = achado & "Setup desalinhado: termina " & Format$(dif, "0") & " min em rela" & ChrW(231) & ChrW(227) & "o ao in" & ChrW(237) & "cio da produ" & ChrW(231) & ChrW(227) & "o; "
                If Not APS_Ig(ops(p).maquina, ops(i).maquina) Then achado = achado & "Setup em outra m" & ChrW(225) & "quina (" & ops(p).maquina & "); "
                If StrComp(ops(p).Produto, ops(i).Produto, vbTextCompare) <> 0 Then achado = achado & "Setup com outro produto; "
                If StrComp(ops(p).lote, ops(i).lote, vbTextCompare) <> 0 Then achado = achado & "Setup com outro lote; "
            End If
        Else
            nAux = nAux + 1
            If Len(Trim$(ops(i).idPai)) = 0 Then
                achado = achado & "formato antigo (sem idPai: n" & ChrW(227) & "o acompanha produ" & ChrW(231) & ChrW(227) & "o); "
            Else
                achou = False
                For j = 1 To n
                    If Len(ops(j).tipoAux) = 0 Then
                        If APS_Ig(ops(j).id, ops(i).idPai) Then achou = True: Exit For
                    End If
                Next j
                If Not achou Then achado = achado & "ORF" & ChrW(195) & "O: produ" & ChrW(231) & ChrW(227) & "o '" & ops(i).idPai & "' n" & ChrW(227) & "o encontrada; "
            End If
            If cPla > 0 And ops(i).linha > 0 Then
                If Len(Trim$(CStr(wsO.Cells(ops(i).linha, cPla).Value))) > 0 Then achado = achado & "Setup/Limpeza com status gravado; "
            End If
        End If

        If UCase$(Left$(ops(i).id, 1)) = "R" And Not IsNumeric(ops(i).id) Then
            achado = achado & "ID provis" & ChrW(243) & "rio (coluna ID vazia nesta linha); "
            nProv = nProv + 1
        End If

        If APS_Ativa(ops(i)) Then
            For j = i + 1 To n
                If APS_Ativa(ops(j)) Then
                    If APS_Ig(ops(i).maquina, ops(j).maquina) Then
                        If APS_Sobrepoe(ops(i), ops(j)) Then achado = achado & "sobrep" & ChrW(245) & "e a linha " & ops(j).linha & "; "
                    End If
                End If
            Next j
            If InStr(nomes, "|OPCARD_" & ops(i).id & "|") = 0 Then achado = achado & "sem card no planejamento (fora do per" & ChrW(237) & "odo exibido ou n" & ChrW(227) & "o desenhado); "
        End If

        r = r + 1
        If Len(ops(i).tipoAux) = 0 Then wsA.Cells(r, 1).Value = "Produ" & ChrW(231) & ChrW(227) & "o" Else wsA.Cells(r, 1).Value = ops(i).tipoAux
        wsA.Cells(r, 2).Value = Abs(ops(i).linha)
        wsA.Cells(r, 3).Value = "'" & ops(i).id
        wsA.Cells(r, 4).Value = ops(i).Produto
        wsA.Cells(r, 5).Value = "'" & ops(i).lote
        wsA.Cells(r, 6).Value = ops(i).maquina
        wsA.Cells(r, 7).Value = ops(i).ini
        wsA.Cells(r, 8).Value = ops(i).Fim
        wsA.Cells(r, 7).NumberFormat = "dd/mm/yyyy hh:mm"
        wsA.Cells(r, 8).NumberFormat = "dd/mm/yyyy hh:mm"
        wsA.Cells(r, 9).Value = Round(ops(i).dur * 1440#, 0)
        wsA.Cells(r, 10).Value = ops(i).status
        wsA.Cells(r, 11).Value = "'" & ops(i).idPai
        wsA.Cells(r, 13).Value = ops(i).ordem
        wsA.Cells(r, 13).NumberFormat = "@"
        If Len(achado) = 0 Then
            wsA.Cells(r, 12).Value = "OK" & IIf(Len(info) > 0, " | INFO: " & info, "")
        Else
            wsA.Cells(r, 12).Value = achado & IIf(Len(info) > 0, " | INFO: " & info, "")
            nProb = nProb + 1
        End If
    Next i

    ' cards do planejamento sem operacao correspondente
    For Each shp In wsP.Shapes
        If Left$(shp.Name, 7) = "OPCARD_" Then
            If InStr(ids, "|" & shp.Name & "|") = 0 Then
                nSob = nSob + 1
                r = r + 1
                wsA.Cells(r, 1).Value = "CARD"
                wsA.Cells(r, 3).Value = "'" & mID$(shp.Name, 8)
                wsA.Cells(r, 12).Value = "CARD SOBRANDO: " & shp.Name & " n" & ChrW(227) & "o tem opera" & ChrW(231) & ChrW(227) & "o em 02_Operacoes"
                nProb = nProb + 1
            End If
        End If
    Next shp

    wsA.Cells(1, 1).Value = "AUDITORIA DO V" & ChrW(205) & "NCULO PRODU" & ChrW(199) & ChrW(195) & "O x SETUP/LIMPEZA (n" & ChrW(227) & "o altera nada)"
    wsA.Cells(1, 1).Font.Bold = True
    cPro = APS_Col(wsO, H_PRO()): cOrd = APS_Col(wsO, H_ORD())
    wsA.Cells(3, 1).Value = "Coluna OP logo depois de Produto: " & IIf(cOrd > 0 And cOrd = cPro + 1, "SIM", "N" & ChrW(195) & "O") & _
                            "  |  OP " & ChrW(233) & " independente do ID t" & ChrW(233) & "cnico (coluna ID oculta)"
    wsA.Cells(2, 1).Value = "Produ" & ChrW(231) & ChrW(245) & "es: " & nProd & "  |  Setup/Limpeza: " & nAux & "  |  IDs provis" & ChrW(243) & "rios: " & nProv & _
                            "  |  Linhas com achado: " & nProb & "  |  Cards sobrando: " & nSob
    wsA.Columns("A:K").AutoFit
    wsA.Columns(12).ColumnWidth = 90
    wsA.Activate

    MsgBox "Auditoria conclu" & ChrW(237) & "da (aba APS_AUDITORIA)." & vbCrLf & vbCrLf & _
           "Produ" & ChrW(231) & ChrW(245) & "es: " & nProd & vbCrLf & "Setup/Limpeza: " & nAux & vbCrLf & _
           "IDs provis" & ChrW(243) & "rios: " & nProv & vbCrLf & "Linhas com achado: " & nProb & vbCrLf & "Cards sobrando: " & nSob, _
           vbInformation, APS_TITULO
    Exit Sub

Falha:
    Application.DisplayAlerts = True
    MsgBox "Erro na auditoria: " & Err.Description, vbExclamation, APS_TITULO
End Sub
