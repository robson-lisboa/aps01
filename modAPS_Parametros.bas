Attribute VB_Name = "modAPS_Parametros"
Option Explicit

'==========================================================
' APS PURAN - CADASTROS E PARAMETROS (dados e regras)
'
' Cada parametro e um REGISTRO com identidade propria:  id | tipo | maquina | codigo do produto | valor | unidade
' Gravados em nome oculto do proprio arquivo (mesmo mecanismo do calendario/OEE: APS_CfgLer / APS_CfgGravar), portanto
' persistem ao salvar/fechar/reabrir. NENHUM valor e fixado no codigo: lotes, conteudo por caixa, velocidades, pesos,
' capacidades etc. sao cadastrados pelo usuario na tela "Cadastros e Parametros" (frmAPS_Parametros).
'
'   Tipo  Parametro                    Escopo                     Unidade
'   1     Tamanho do lote              produto x maquina          Comprimidos
'   2     Conteudo por caixa           produto x maquina          Comprimidos/caixa
'   3     Quantidade de comprimidos    produto                    Comprimidos
'   4     Peso                         produto                    mg / g / kg
'   5     Granulacao                   produto                    texto livre
'   6     Sala de mistura              produto                    Sala
'   7     Velocidade nominal           produto x maquina          Caixas/minuto ou Comprimidos/minuto
'   8     Capacidade                   etapa (Sala de Mistura / Granulacao)   kg
'   9     Capacidade meia dosagem      Sala de Mistura x produto  kg   (lista explicita de produtos de meia dosagem)
' Tipos 3, 4, 5, 6 e 8 sao cadastro de referencia: nenhuma formula os usa (a duracao de Sala/Granulacao segue manual).
'
' Este modulo NAO cria operacoes nem cards e NAO altera operacoes ja gravadas.
' Quantidade oficial do lote = COMPRIMIDOS. Caixas = comprimidos / conteudo por caixa (so para o calculo do tempo).
'==========================================================

Private Const K_REG As String = "PAR_REG"

Public Const PAR_LOTE As Long = 1
Public Const PAR_CPC As Long = 2
Public Const PAR_QTDCOMP As Long = 3
Public Const PAR_PESO As Long = 4
Public Const PAR_GRANUL As Long = 5
Public Const PAR_SALA As Long = 6
Public Const PAR_VEL As Long = 7
Public Const PAR_CAPAC As Long = 8
Public Const PAR_MEIA As Long = 9
Public Const PAR_NTIPOS As Long = 9

Public Const UN_CX_MIN As String = "Caixas/minuto"
Public Const UN_COMP_MIN As String = "Comprimidos/minuto"

'----------------------------------------------------------
' Catalogo dos tipos (so regras de cadastro; nenhum valor de producao)
'----------------------------------------------------------
Public Function APS_RegNomeTipo(ByVal t As Long) As String
    Select Case t
        Case PAR_LOTE: APS_RegNomeTipo = "Tamanho do lote"
        Case PAR_CPC: APS_RegNomeTipo = "Conte" & ChrW(250) & "do por caixa"
        Case PAR_QTDCOMP: APS_RegNomeTipo = "Quantidade de comprimidos"
        Case PAR_PESO: APS_RegNomeTipo = "Peso"
        Case PAR_GRANUL: APS_RegNomeTipo = "Granula" & ChrW(231) & ChrW(227) & "o"
        Case PAR_SALA: APS_RegNomeTipo = "Sala de mistura"
        Case PAR_VEL: APS_RegNomeTipo = "Velocidade nominal"
        Case PAR_CAPAC: APS_RegNomeTipo = "Capacidade"
        Case PAR_MEIA: APS_RegNomeTipo = "Capacidade meia dosagem"
    End Select
End Function

' Escopo: "PM" produto x maquina, "P" produto, "E" etapa, "EP" etapa x produto
Public Function APS_RegEscopo(ByVal t As Long) As String
    Select Case t
        Case PAR_LOTE, PAR_CPC, PAR_VEL: APS_RegEscopo = "PM"
        Case PAR_QTDCOMP, PAR_PESO, PAR_GRANUL, PAR_SALA: APS_RegEscopo = "P"
        Case PAR_CAPAC: APS_RegEscopo = "E"
        Case PAR_MEIA: APS_RegEscopo = "EP"
    End Select
End Function

' Grupo na tela: "B" = parametros do produto, "C" = parametros da maquina ou etapa
Public Function APS_RegGrupo(ByVal t As Long) As String
    Select Case t
        Case PAR_LOTE, PAR_CPC, PAR_QTDCOMP, PAR_PESO, PAR_GRANUL: APS_RegGrupo = "B"
        Case Else: APS_RegGrupo = "C"
    End Select
End Function

' "I" inteiro > 0, "N" numero > 0, "T" texto
Public Function APS_RegKindValor(ByVal t As Long) As String
    Select Case t
        Case PAR_LOTE, PAR_QTDCOMP: APS_RegKindValor = "I"
        Case PAR_GRANUL, PAR_SALA: APS_RegKindValor = "T"
        Case Else: APS_RegKindValor = "N"
    End Select
End Function

' Unidades do tipo. Devolve 0 = unidade de texto livre; 1 = unidade fixa; 2 ou mais = escolher.
Public Function APS_RegUnidades(ByVal t As Long, ByRef u() As String) As Long
    ReDim u(1 To 3)
    Select Case t
        Case PAR_LOTE, PAR_QTDCOMP: u(1) = "Comprimidos": APS_RegUnidades = 1
        Case PAR_CPC: u(1) = "Comprimidos/caixa": APS_RegUnidades = 1
        Case PAR_PESO: u(1) = "mg": u(2) = "g": u(3) = "kg": APS_RegUnidades = 3
        Case PAR_GRANUL: u(1) = "Texto": APS_RegUnidades = 1
        Case PAR_SALA: u(1) = "Sala": APS_RegUnidades = 1
        Case PAR_VEL: u(1) = UN_CX_MIN: u(2) = UN_COMP_MIN: APS_RegUnidades = 2
        Case PAR_CAPAC, PAR_MEIA: u(1) = "kg": APS_RegUnidades = 1
    End Select
End Function

' Maquinas / etapas aceitas pelo tipo
Public Function APS_RegMaquinas(ByVal t As Long, ByRef m() As String) As Long
    ReDim m(1 To 4)
    Select Case t
        Case PAR_LOTE, PAR_CPC
            m(1) = "FETTE 2090": m(2) = "FETTE 2": m(3) = "Mediseal": m(4) = "Blisterflex": APS_RegMaquinas = 4
        Case PAR_VEL
            m(1) = "FETTE 2090": m(2) = "FETTE 2": m(3) = "Mediseal": APS_RegMaquinas = 3
        Case PAR_CAPAC
            m(1) = "Sala de Mistura": m(2) = "Granula" & ChrW(231) & ChrW(227) & "o": APS_RegMaquinas = 2
        Case PAR_MEIA
            m(1) = "Sala de Mistura": APS_RegMaquinas = 1
    End Select
End Function

' Todas as maquinas/etapas que aparecem em algum cadastro (para o filtro da tela)
Public Function APS_RegTodasMaquinas(ByRef m() As String) As Long
    ReDim m(1 To 6)
    m(1) = "FETTE 2090"
    m(2) = "FETTE 2"
    m(3) = "Mediseal"
    m(4) = "Blisterflex"
    m(5) = "Sala de Mistura"
    m(6) = "Granula" & ChrW(231) & ChrW(227) & "o"
    APS_RegTodasMaquinas = 6
End Function

Public Function APS_ParUsaVelocidade(ByVal maq As String) As Boolean
    APS_ParUsaVelocidade = APS_EhFette(maq) Or APS_EhMediseal(maq)
End Function


'----------------------------------------------------------
' Numeros digitados / exibidos (independente da configuracao regional do Windows)
'   aceita "4.000.000", "1.000,5", "12,5", "0.5", "12.5"; ponto com 3 digitos apos o primeiro grupo = milhar.
'   Devolve 0 se invalido. (APS_ParaQtd da Base tratava "0.5" como 5.)
'----------------------------------------------------------
Public Function APS_ParNumero(ByVal s As String) As Double
    Dim t As String, p() As String, i As Long, ok As Boolean

    t = Replace(Trim$(s), " ", "")
    If Len(t) = 0 Then Exit Function
    If t Like "*[!0-9.,]*" Then Exit Function

    If InStr(t, ",") > 0 Then
        If InStr(t, ",") <> InStrRev(t, ",") Then Exit Function
        If InStr(InStr(t, ",") + 1, t, ".") > 0 Then Exit Function
        t = Replace(t, ".", "")
        t = Replace(t, ",", ".")
    ElseIf InStr(t, ".") > 0 Then
        p = Split(t, ".")
        ok = (Len(p(0)) >= 1 And Len(p(0)) <= 3)
        If ok Then
            If Left$(p(0), 1) = "0" Then ok = False
        End If
        For i = 1 To UBound(p)
            If Len(p(i)) <> 3 Then ok = False
        Next i
        If ok Then
            t = Replace(t, ".", "")
        ElseIf UBound(p) <> 1 Then
            Exit Function
        End If
    End If

    If t = "." Then Exit Function
    APS_ParNumero = Val(t)
End Function

' Texto do valor guardado (sempre com ponto decimal, sem separador de milhar)
Public Function APS_ParTexto(ByVal num As Double) As String
    APS_ParTexto = Trim$(Str$(num))
End Function

' Exibicao: milhar com ponto, decimal com virgula (ex.: 4.000.000 e 12,5)
Public Function APS_ParFmtNum(ByVal v As String) As String
    Dim num As Double, ip As String, fp As String, r As String, i As Long, p As Long, t As String

    num = Val(v)
    t = Trim$(Str$(num))
    p = InStr(t, ".")

    If p > 0 Then
        ip = Left$(t, p - 1)
        fp = Mid$(t, p + 1)
    Else
        ip = t
    End If

    For i = Len(ip) To 1 Step -1
        r = Mid$(ip, i, 1) & r
        If (Len(ip) - i + 1) Mod 3 = 0 And i > 1 Then
            If Mid$(ip, i - 1, 1) <> "-" Then r = "." & r
        End If
    Next i

    If Len(fp) > 0 Then r = r & "," & fp
    APS_ParFmtNum = r
End Function

'----------------------------------------------------------
' Pesquisa de produtos ja cadastrados (codigo OU nome, sem diferenciar maiusculas). Nao cria nem altera nada.
' Texto vazio = todos os produtos. Devolve a quantidade de resultados.
'----------------------------------------------------------
Public Function APS_ParBuscaProduto(ByVal q As String, ByRef cods() As String, ByRef nomes() As String) As Long
    Dim c() As String, n() As String, l() As Double, v() As Double
    Dim k As Long, i As Long, j As Long, ok As Boolean

    k = APS_Produtos(c, n, l, v)
    q = Trim$(q)

    ReDim cods(1 To k + 1)
    ReDim nomes(1 To k + 1)

    For i = 1 To k
        If Len(q) = 0 Then
            ok = True
        Else
            ok = (InStr(1, c(i), q, vbTextCompare) > 0) Or (InStr(1, n(i), q, vbTextCompare) > 0)
        End If

        If ok Then
            j = j + 1
            cods(j) = c(i)
            nomes(j) = n(i)
        End If
    Next i

    APS_ParBuscaProduto = j
End Function

'----------------------------------------------------------
' Leitura / gravacao (identidade do registro = id; chave unica = tipo + maquina + produto)
'----------------------------------------------------------
Private Function Limpa(ByVal s As String) As String
    s = Replace(s, ";", " ")
    s = Replace(s, "|", " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Limpa = Trim$(s)
End Function

Public Function APS_RegListar(ByRef id() As Long, ByRef tp() As Long, ByRef mq() As String, _
                              ByRef cd() As String, ByRef vl() As String, ByRef un() As String) As Long
    Dim s As String, a() As String, f() As String, i As Long, n As Long

    ReDim id(1 To 1)
    ReDim tp(1 To 1)
    ReDim mq(1 To 1)
    ReDim cd(1 To 1)
    ReDim vl(1 To 1)
    ReDim un(1 To 1)

    s = APS_CfgLer(K_REG, "")
    If Len(s) = 0 Then Exit Function

    a = Split(s, ";")

    ReDim id(1 To UBound(a) + 1)
    ReDim tp(1 To UBound(a) + 1)
    ReDim mq(1 To UBound(a) + 1)
    ReDim cd(1 To UBound(a) + 1)
    ReDim vl(1 To UBound(a) + 1)
    ReDim un(1 To UBound(a) + 1)

    For i = 0 To UBound(a)
        f = Split(a(i), "|")
        If UBound(f) >= 5 Then
            If Val(f(0)) > 0 And Val(f(1)) >= 1 And Val(f(1)) <= PAR_NTIPOS Then
                n = n + 1
                id(n) = CLng(Val(f(0)))
                tp(n) = CLng(Val(f(1)))
                mq(n) = f(2)
                cd(n) = f(3)
                vl(n) = f(4)
                un(n) = f(5)
            End If
        End If
    Next i

    APS_RegListar = n
End Function

Private Sub GravarTudo(ByVal n As Long, ByRef id() As Long, ByRef tp() As Long, ByRef mq() As String, _
                       ByRef cd() As String, ByRef vl() As String, ByRef un() As String)
    Dim s As String, i As Long

    For i = 1 To n
        If Len(s) > 0 Then s = s & ";"
        s = s & CStr(id(i)) & "|" & CStr(tp(i)) & "|" & mq(i) & "|" & cd(i) & "|" & vl(i) & "|" & un(i)
    Next i

    APS_CfgGravar K_REG, s
End Sub

Private Function AcharChave(ByVal n As Long, ByRef tp() As Long, ByRef mq() As String, ByRef cd() As String, _
                            ByVal t As Long, ByVal m As String, ByVal c As String) As Long
    Dim i As Long

    For i = 1 To n
        If tp(i) = t And APS_Ig(mq(i), m) And APS_Ig(cd(i), c) Then
            AcharChave = i
            Exit Function
        End If
    Next i
End Function

Private Function AcharId(ByVal n As Long, ByRef id() As Long, ByVal x As Long) As Long
    Dim i As Long

    For i = 1 To n
        If id(i) = x Then
            AcharId = i
            Exit Function
        End If
    Next i
End Function

' Busca o valor cadastrado (texto invariante) e a unidade da combinacao. False = nao cadastrado.
Public Function APS_RegBuscar(ByVal t As Long, ByVal maq As String, ByVal cod As String, _
                              ByRef valor As String, ByRef unid As String) As Boolean
    Dim id() As Long, tp() As Long, mq() As String, cd() As String, vl() As String, un() As String
    Dim n As Long, i As Long

    valor = ""
    unid = ""
    n = APS_RegListar(id, tp, mq, cd, vl, un)
    i = AcharChave(n, tp, mq, cd, t, maq, cod)

    If i > 0 Then
        valor = vl(i)
        unid = un(i)
        APS_RegBuscar = True
    End If
End Function

Private Function NumDe(ByVal t As Long, ByVal maq As String, ByVal cod As String) As Double
    Dim v As String, u As String

    If APS_RegBuscar(t, maq, cod, v, u) Then NumDe = Val(v)
End Function

' Valida e NORMALIZA os campos (maq, cod, valor, unid voltam no formato gravado). Nada e gravado. False = msg explica.
Public Function APS_RegValidar(ByVal t As Long, ByRef maq As String, ByRef cod As String, _
                               ByRef valor As String, ByRef unid As String, ByRef msg As String) As Boolean
    Dim esc As String, ms() As String, k As Long, j As Long, ok As Boolean
    Dim us() As String, ku As Long, kind As String, num As Double

    If t < 1 Or t > PAR_NTIPOS Then msg = "Selecione o par" & ChrW(226) & "metro.": Exit Function

    esc = APS_RegEscopo(t)
    cod = Trim$(cod)
    maq = Trim$(maq)

    If esc = "PM" Or esc = "P" Or esc = "EP" Then
        If Len(cod) = 0 Then msg = "Selecione o produto.": Exit Function
        If APS_IndiceProduto(cod) = 0 Then msg = "Produto inexistente no cadastro de produtos.": Exit Function
        cod = Trim$(cod)
    Else
        cod = ""
    End If

    If esc = "PM" Or esc = "E" Or esc = "EP" Then
        If Len(maq) = 0 Then msg = "Selecione a m" & ChrW(225) & "quina ou etapa.": Exit Function
        k = APS_RegMaquinas(t, ms)
        For j = 1 To k
            If APS_Ig(ms(j), maq) Then ok = True: maq = ms(j): Exit For
        Next j
        If Not ok Then msg = "Este par" & ChrW(226) & "metro n" & ChrW(227) & "o se aplica " & ChrW(224) & " m" & ChrW(225) & "quina informada.": Exit Function
    Else
        maq = ""
    End If

    ' valor
    kind = APS_RegKindValor(t)
    valor = Trim$(valor)

    If Len(valor) = 0 Then msg = "Informe o valor.": Exit Function

    If kind = "T" Then
        valor = Limpa(valor)
        If Len(valor) = 0 Or Len(valor) > 40 Then msg = "Texto inv" & ChrW(225) & "lido (at" & ChrW(233) & " 40 caracteres).": Exit Function
    Else
        num = APS_ParNumero(valor)
        If num <= 0 Then msg = "Valor inv" & ChrW(225) & "lido: use um n" & ChrW(250) & "mero maior que zero.": Exit Function
        If kind = "I" And num <> Int(num) Then msg = "Use um n" & ChrW(250) & "mero inteiro de comprimidos.": Exit Function
        valor = APS_ParTexto(num)
    End If

    ' unidade
    ku = APS_RegUnidades(t, us)
    If ku = 0 Then
        unid = Limpa(unid)
        If Len(unid) = 0 Or Len(unid) > 20 Then msg = "Informe a unidade (at" & ChrW(233) & " 20 caracteres).": Exit Function
    Else
        ok = False
        For j = 1 To ku
            If APS_Ig(us(j), unid) Then ok = True: unid = us(j): Exit For
        Next j
        If Not ok Then
            If ku = 1 Then
                unid = us(1)
            Else
                msg = "Selecione a unidade.": Exit Function
            End If
        End If
    End If


    APS_RegValidar = True
End Function

' Valida e grava. id = 0 -> novo; id > 0 -> edita (so valor e unidade; tipo, maquina e produto sao a identidade).
' Devolve o id do registro, ou 0 com a mensagem em msg.
Public Function APS_RegSalvar(ByVal idReg As Long, ByVal t As Long, ByVal maq As String, ByVal cod As String, _
                              ByVal valor As String, ByVal unid As String, ByRef msg As String) As Long
    Dim id() As Long, tp() As Long, mq() As String, cd() As String, vl() As String, un() As String
    Dim n As Long, i As Long, j As Long, mx As Long

    If Not APS_RegValidar(t, maq, cod, valor, unid, msg) Then Exit Function

    n = APS_RegListar(id, tp, mq, cd, vl, un)

    If idReg > 0 Then
        i = AcharId(n, id, idReg)
        If i = 0 Then msg = "O registro n" & ChrW(227) & "o existe mais (foi removido).": Exit Function
        If tp(i) <> t Or Not APS_Ig(mq(i), maq) Or Not APS_Ig(cd(i), cod) Then
            msg = "Tipo, m" & ChrW(225) & "quina e produto n" & ChrW(227) & "o podem ser alterados na edi" & ChrW(231) & ChrW(227) & "o. Remova e cadastre novamente.": Exit Function
        End If
    Else
        If AcharChave(n, tp, mq, cd, t, maq, cod) > 0 Then
            msg = "J" & ChrW(225) & " existe este par" & ChrW(226) & "metro para a mesma combina" & ChrW(231) & ChrW(227) & "o. Use Editar.": Exit Function
        End If
        For j = 1 To n
            If id(j) > mx Then mx = id(j)
        Next j
        n = n + 1
        ReDim Preserve id(1 To n)
        ReDim Preserve tp(1 To n)
        ReDim Preserve mq(1 To n)
        ReDim Preserve cd(1 To n)
        ReDim Preserve vl(1 To n)
        ReDim Preserve un(1 To n)
        i = n
        id(i) = mx + 1
        tp(i) = t
        mq(i) = maq
        cd(i) = cod
    End If

    vl(i) = valor
    un(i) = unid

    GravarTudo n, id, tp, mq, cd, vl, un
    APS_RegSalvar = id(i)
End Function

' Remove definitivamente o registro (a tela pede a confirmacao). Nada e recriado automaticamente.
Public Function APS_RegRemover(ByVal idReg As Long) As Boolean
    Dim id() As Long, tp() As Long, mq() As String, cd() As String, vl() As String, un() As String
    Dim n As Long, i As Long, j As Long

    n = APS_RegListar(id, tp, mq, cd, vl, un)
    i = AcharId(n, id, idReg)
    If i = 0 Then Exit Function

    For j = i To n - 1
        id(j) = id(j + 1)
        tp(j) = tp(j + 1)
        mq(j) = mq(j + 1)
        cd(j) = cd(j + 1)
        vl(j) = vl(j + 1)
        un(j) = un(j + 1)
    Next j

    GravarTudo n - 1, id, tp, mq, cd, vl, un
    APS_RegRemover = True
End Function

' Grava de uma vez os parametros de UM produto (e/ou capacidades gerais) vindos da tela. Cada item:  tipo, maquina, produto, valor, unidade.
'   valor preenchido -> cria ou atualiza o registro da combinacao (tipo + maquina + produto): nunca duplica
'   valor vazio      -> remove o registro dessa combinacao, se existir
' Tudo e validado ANTES de gravar (um erro nao grava nada). Depois de gravar, o cadastro e RELIDO e conferido item a item;
' se algo nao confere, o estado anterior e restaurado e a funcao devolve False.
Public Function APS_RegAplicar(ByVal n As Long, ByRef t() As Long, ByRef m() As String, ByRef c() As String, _
                               ByRef v() As String, ByRef u() As String, ByRef msg As String, _
                               ByRef nGrav As Long, ByRef nRem As Long) As Boolean
    Dim id() As Long, tp() As Long, mq() As String, cd() As String, vl() As String, un() As String
    Dim nn As Long, i As Long, j As Long, k As Long, mx As Long, antes As String, m2 As String
    Dim nm() As String, nc() As String, nv() As String, nu() As String, falha As Boolean
    Dim id2() As Long, tp2() As Long, mq2() As String, cd2() As String, vl2() As String, un2() As String, n2 As Long

    nGrav = 0
    nRem = 0
    If n < 1 Then Exit Function

    ReDim nm(1 To n)
    ReDim nc(1 To n)
    ReDim nv(1 To n)
    ReDim nu(1 To n)

    ' 1) validar tudo, sem gravar
    For i = 1 To n
        nm(i) = m(i)
        nc(i) = c(i)
        nv(i) = Trim$(v(i))
        nu(i) = u(i)

        If Len(nv(i)) > 0 Then
            If Not APS_RegValidar(t(i), nm(i), nc(i), nv(i), nu(i), m2) Then
                msg = APS_RegNomeTipo(t(i))
                If Len(m(i)) > 0 Then msg = msg & " (" & m(i) & ")"
                msg = msg & ": " & m2
                Exit Function
            End If
        End If
    Next i

    ' 2) aplicar em memoria
    antes = APS_CfgLer(K_REG, "")
    nn = APS_RegListar(id, tp, mq, cd, vl, un)

    For i = 1 To nn
        If id(i) > mx Then mx = id(i)
    Next i

    For i = 1 To n
        k = AcharChave(nn, tp, mq, cd, t(i), nm(i), nc(i))

        If Len(nv(i)) = 0 Then
            If k > 0 Then
                For j = k To nn - 1
                    id(j) = id(j + 1)
                    tp(j) = tp(j + 1)
                    mq(j) = mq(j + 1)
                    cd(j) = cd(j + 1)
                    vl(j) = vl(j + 1)
                    un(j) = un(j + 1)
                Next j
                nn = nn - 1
                nRem = nRem + 1
            End If
        ElseIf k > 0 Then
            If vl(k) <> nv(i) Or un(k) <> nu(i) Then nGrav = nGrav + 1
            vl(k) = nv(i)
            un(k) = nu(i)
        Else
            nn = nn + 1
            ReDim Preserve id(1 To nn)
            ReDim Preserve tp(1 To nn)
            ReDim Preserve mq(1 To nn)
            ReDim Preserve cd(1 To nn)
            ReDim Preserve vl(1 To nn)
            ReDim Preserve un(1 To nn)
            mx = mx + 1
            id(nn) = mx
            tp(nn) = t(i)
            mq(nn) = nm(i)
            cd(nn) = nc(i)
            vl(nn) = nv(i)
            un(nn) = nu(i)
            nGrav = nGrav + 1
        End If
    Next i

    ' 3) gravar
    If nn = 0 Then
        APS_CfgGravar K_REG, ""
    Else
        GravarTudo nn, id, tp, mq, cd, vl, un
    End If

    ' 4) reler do cadastro gravado e conferir cada item
    n2 = APS_RegListar(id2, tp2, mq2, cd2, vl2, un2)
    falha = False

    For i = 1 To n
        k = AcharChave(n2, tp2, mq2, cd2, t(i), nm(i), nc(i))

        If Len(nv(i)) = 0 Then
            If k > 0 Then falha = True
        ElseIf k = 0 Then
            falha = True
        ElseIf vl2(k) <> nv(i) Or un2(k) <> nu(i) Then
            falha = True
        End If
    Next i

    If falha Then
        APS_CfgGravar K_REG, antes
        msg = "O cadastro gravado n" & ChrW(227) & "o confere com o informado. O estado anterior foi restaurado."
        nGrav = 0
        nRem = 0
        Exit Function
    End If

    APS_RegAplicar = True
End Function

'----------------------------------------------------------
' Consultas usadas pela tela da operacao. 0 = nao cadastrado (NUNCA assume valor).
'----------------------------------------------------------
Public Function APS_ParLote(ByVal maq As String, ByVal cod As String) As Double
    APS_ParLote = NumDe(PAR_LOTE, maq, cod)
End Function

Public Function APS_ParCPC(ByVal maq As String, ByVal cod As String) As Double
    APS_ParCPC = NumDe(PAR_CPC, maq, cod)
End Function

Public Function APS_ParVel(ByVal maq As String, ByVal cod As String) As Double
    If Not APS_ParUsaVelocidade(maq) Then Exit Function
    APS_ParVel = NumDe(PAR_VEL, maq, cod)
End Function

Public Function APS_ParVelUnid(ByVal maq As String, ByVal cod As String) As String
    Dim v As String, u As String

    If APS_RegBuscar(PAR_VEL, maq, cod, v, u) Then APS_ParVelUnid = u
End Function

' Duracao calculada pelo parametro?  Mediseal: sempre automatica (velocidade cadastrada OU a velocidade-base ja existente).
' FETTE: automatica SO quando ha velocidade cadastrada.
Public Function APS_DurAutomatica(ByVal maq As String, ByVal cod As String) As Boolean
    If APS_EhMediseal(maq) Then
        APS_DurAutomatica = True
    ElseIf APS_EhFette(maq) Then
        APS_DurAutomatica = (APS_ParVel(maq, cod) > 0)
    End If
End Function

'----------------------------------------------------------
' Capacidades (kg) - Sala de Mistura e Granulacao (cadastro de REFERENCIA; a duracao dessas etapas segue manual).
' Excecao de meia dosagem: SOMENTE na Sala de Mistura e SOMENTE para produtos com registro de meia dosagem.
'----------------------------------------------------------
Public Function APS_EhMeiaDosagem(ByVal cod As String) As Boolean
    Dim v As String, u As String

    APS_EhMeiaDosagem = APS_RegBuscar(PAR_MEIA, "Sala de Mistura", cod, v, u)
End Function

Public Function APS_CapacidadeKg(ByVal maq As String, ByVal cod As String) As Double
    Dim v As String, u As String, nm As String

    If APS_EhSala(maq) Then
        nm = "Sala de Mistura"
        If APS_RegBuscar(PAR_MEIA, nm, cod, v, u) Then
            APS_CapacidadeKg = Val(v)
            Exit Function
        End If
    ElseIf APS_EhGranulacao(maq) Then
        nm = "Granula" & ChrW(231) & ChrW(227) & "o"
    Else
        Exit Function
    End If

    If APS_RegBuscar(PAR_CAPAC, nm, "", v, u) Then APS_CapacidadeKg = Val(v)
End Function

'----------------------------------------------------------
' Calculo do tempo (FETTE / Mediseal com velocidade cadastrada)
'   caixas        = comprimidos / conteudo por caixa
'   tempo nominal = caixas / velocidade nominal (caixas/min)         <- calculo de REFERENCIA: sem OEE, setup nem paradas
'                   (velocidade em comprimidos/min: comprimidos / velocidade)
'   tempo com OEE = tempo nominal / OEE                              <- operacional (OEE da operacao, aplicado uma vez;
'                   OEE 100% = tempo nominal). Setup e paradas continuam sendo tratados pelo motor/calendario.
' O lote oficial (comprimidos) nunca e alterado pelo calculo. Parametro ausente/invalido = sem calculo.
'----------------------------------------------------------
Public Function APS_CalcTempo(ByVal comp As Double, ByVal cpc As Double, ByVal vel As Double, _
                              ByVal porCaixa As Boolean, ByVal oee As Double, _
                              ByRef caixas As Double, ByRef segNom As Double, ByRef segOee As Double) As Boolean
    caixas = 0
    segNom = 0
    segOee = 0

    If comp <= 0 Or vel <= 0 Or oee <= 0 Or oee > 1.0000001 Then Exit Function

    If porCaixa Then
        If cpc <= 0 Then Exit Function
        caixas = comp / cpc
        segNom = caixas / vel * 60#
    Else
        If cpc > 0 Then caixas = comp / cpc
        segNom = comp / vel * 60#
    End If

    segOee = segNom / oee
    APS_CalcTempo = (segNom > 0)
End Function

' Segundos -> "h h mm min ss s"
Public Function APS_FmtHMS(ByVal segundos As Double) As String
    Dim t As Long

    t = CLng(Round(segundos, 0))
    APS_FmtHMS = CStr(t \ 3600) & "h " & Format$(((t Mod 3600) \ 60), "00") & "min " & Format$((t Mod 60), "00") & "s"
End Function

'----------------------------------------------------------
' Resumo do cadastro para a Nova Operacao (produto + maquina). Diz claramente o que existe e o que falta.
'----------------------------------------------------------
Public Function APS_ResumoCadastro(ByVal maq As String, ByVal cod As String, ByRef faltando As Boolean) As String
    Dim lp As Double, cpc As Double, vel As Double, cap As Double, s As String

    faltando = False
    If Len(maq) = 0 Or Len(cod) = 0 Then Exit Function

    If APS_EhSala(maq) Or APS_EhGranulacao(maq) Then
        cap = APS_CapacidadeKg(maq, cod)
        If cap > 0 Then
            s = "Capacidade cadastrada: " & APS_ParFmtNum(APS_ParTexto(cap)) & " kg (refer" & ChrW(234) & "ncia). Dura" & ChrW(231) & ChrW(227) & "o manual."
        Else
            s = "Capacidade n" & ChrW(227) & "o cadastrada para esta etapa. Dura" & ChrW(231) & ChrW(227) & "o manual."
            faltando = True
        End If
        APS_ResumoCadastro = s
        Exit Function
    End If

    lp = APS_ParLote(maq, cod)
    cpc = APS_ParCPC(maq, cod)
    vel = APS_ParVel(maq, cod)

    If lp <= 0 And cpc <= 0 And vel <= 0 Then
        APS_ResumoCadastro = "Sem cadastro de par" & ChrW(226) & "metros para este produto nesta m" & ChrW(225) & "quina (Cadastros e Par" & ChrW(226) & "metros). Informe quantidade e dura" & ChrW(231) & ChrW(227) & "o."
        faltando = True
        Exit Function
    End If

    If lp > 0 Then
        s = "Lote padr" & ChrW(227) & "o: " & APS_ParFmtNum(APS_ParTexto(lp)) & " comprimidos"
    Else
        s = "Lote padr" & ChrW(227) & "o: n" & ChrW(227) & "o cadastrado"
        faltando = True
    End If

    If cpc > 0 Then
        s = s & vbCr & "Comprimidos por caixa: " & APS_ParFmtNum(APS_ParTexto(cpc))
    Else
        s = s & vbCr & "Comprimidos por caixa: n" & ChrW(227) & "o cadastrado"
    End If

    If APS_ParUsaVelocidade(maq) Then
        If vel > 0 Then
            s = s & vbCr & "Velocidade: " & APS_ParFmtNum(APS_ParTexto(vel)) & " " & LCase$(APS_ParVelUnid(maq, cod))
        ElseIf APS_EhMediseal(maq) Then
            s = s & vbCr & "Velocidade: n" & ChrW(227) & "o cadastrada (usa a velocidade-base existente do produto)"
        Else
            s = s & vbCr & "Velocidade: n" & ChrW(227) & "o cadastrada (dura" & ChrW(231) & ChrW(227) & "o manual)"
            faltando = True
        End If
    End If

    APS_ResumoCadastro = s
End Function

'----------------------------------------------------------
' Abrir a tela / botao
'----------------------------------------------------------
Public Sub APS_AbrirParametros()
    Dim f As Object, nomeFrm As String

    nomeFrm = "frmAPS_Parametros"
    On Error GoTo Falha
    Set f = VBA.UserForms.Add(nomeFrm)
    f.Show
    Unload f
    Set f = Nothing
    Exit Sub

Falha:
    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel abrir Cadastros e Par" & ChrW(226) & "metros: " & Err.Description & _
           " (erro " & Err.Number & ")." & vbCrLf & vbCrLf & _
           "Importe o arquivo frmAPS_Parametros.frm (VBE > Arquivo > Importar arquivo).", vbExclamation, APS_TITULO
    On Error Resume Next
    If Not f Is Nothing Then Unload f
End Sub

' Cria (ou recria) o botao em 1_Planejamento SEM cobrir nada: na linha 1, a direita da coluna A (faixa do cabecalho),
' acima do texto da data. Se o cabecalho nao tiver esse espaco, use Alt+F8 > APS_AbrirParametros. Rodar UMA vez.
' Botao CADASTROS E PARAMETROS na tela de planejamento (1_Planejamento), area A1 (congelada: sempre visivel).
' Fica ABAIXO do ultimo botao que ja existe em A1 (NOVA OPERACAO / HOJE / CALENDARIO), mesma largura da coluna,
' mesmo estilo; a linha 1 so cresce o necessario. Nao mexe nos outros botoes. Idempotente (pode rodar de novo).
' Devolve False (e nao cria) se a posicao ainda ficasse sobre outro objeto.
Public Function APS_GarantirBotaoParametros() As Boolean
    Dim ws As Worksheet, est As Boolean, a As Range, shp As Shape, b As Shape
    Dim baixo As Double, topo As Double, alt As Double, nec As Double, nm As String, tl As Range

    On Error GoTo Falha
    Set ws = APS_Aba(APS_ABA_PLAN)
    If ws Is Nothing Then Exit Function
    est = APS_Liberar(ws)
    Set a = ws.Range("A1")
    alt = 18

    ' fim do botao mais baixo ja existente na area A1 (padrao: HOJE/CALENDARIO terminam em 73 pt)
    baixo = a.Top + 73
    For Each shp In ws.Shapes
        nm = shp.Name
        If Left$(nm, 8) = "APS_BTN_" And nm <> "APS_BTN_PAR" And nm <> "APS_BTN_DIAS" Then
            Set tl = Nothing
            On Error Resume Next
            Set tl = shp.TopLeftCell
            On Error GoTo Falha
            If Not tl Is Nothing Then
                If tl.Row = 1 And tl.Column = 1 Then
                    If shp.Top + shp.Height > baixo Then baixo = shp.Top + shp.Height
                End If
            End If
        End If
    Next shp
    topo = baixo + 3
    nec = topo + alt + 6 - a.Top
    If ws.Rows(1).RowHeight < nec Then ws.Rows(1).RowHeight = nec

    On Error Resume Next
    ws.Shapes("APS_BTN_PAR").Delete
    On Error GoTo Falha
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, a.Left + 4, topo, a.Width - 8, alt)
    b.Name = "APS_BTN_PAR"
    b.OnAction = "APS_AbrirParametros"
    b.Fill.ForeColor.RGB = RGB(84, 110, 122)
    b.Line.Visible = msoFalse
    b.Placement = xlMove
    With b.TextFrame2
        .TextRange.Text = "CADASTROS E PAR" & ChrW(194) & "METROS"
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With

    ' confere: nenhum outro botao/objeto da area A1 pode cruzar com este
    APS_GarantirBotaoParametros = True
    For Each shp In ws.Shapes
        If shp.Name <> "APS_BTN_PAR" And Left$(shp.Name, 7) <> "OPCARD_" And shp.Name <> "APS_BTN_DIAS" Then
            If shp.Left < b.Left + b.Width And shp.Left + shp.Width > b.Left And shp.Top < b.Top + b.Height And shp.Top + shp.Height > b.Top Then
                APS_GarantirBotaoParametros = False
                b.Delete
                Exit For
            End If
        End If
    Next shp
    APS_Reproteger ws, est
    Exit Function

Falha:
    APS_GarantirBotaoParametros = False
    On Error Resume Next
    APS_Reproteger ws, est
End Function

' Macro manual (Alt+F8): cria/recria o botao e avisa o resultado.
Public Sub APS_CriarBotaoParametros()
    If APS_GarantirBotaoParametros() Then
        MsgBox "Bot" & ChrW(227) & "o CADASTROS E PAR" & ChrW(194) & "METROS criado na 1_Planejamento, na " & ChrW(225) & "rea A1, abaixo de HOJE e CALEND" & ChrW(193) & "RIO.", vbInformation, APS_TITULO
    Else
        MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel criar o bot" & ChrW(227) & "o sem sobrepor outro objeto. Nada foi criado." & vbCrLf & _
               "Use Alt+F8 > APS_AbrirParametros para abrir a tela.", vbExclamation, APS_TITULO
    End If
End Sub
