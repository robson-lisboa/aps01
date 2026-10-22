Attribute VB_Name = "modAPS_Calendario"
Option Explicit

'==========================================================
' APS PURAN - CALENDARIO DE PRODUCAO
' Dias trabalhados por dia da semana, excecoes por data
' (feriados / dias especiais) e paradas (periodos).
' Configuracao gravada no proprio arquivo (Names ocultos).
'==========================================================

Public Const CAL_SEM_PADRAO As String = "1;S;06:00;22:00|2;S;06:00;22:00|3;S;06:00;22:00|4;S;06:00;22:00|5;S;06:00;22:00|6;N;;|7;N;;"

Private mCarregado As Boolean
Private mTrab(1 To 7) As Boolean
Private mIni(1 To 7) As Double
Private mFim(1 To 7) As Double
Private mNExc As Long
Private mExcData() As Double
Private mExcTrab() As Boolean
Private mExcIni() As Double
Private mExcFim() As Double
Private mNPar As Long
Private mParIni() As Double
Private mParFim() As Double

Public Sub APS_CalRecarregar()
    mCarregado = False
End Sub

'----------------------------------------------------------
' Leitura da configuracao
'----------------------------------------------------------
Private Sub Carregar()
    Dim s As String, itens() As String, p() As String, i As Long, d As Long
    Dim h1 As Double, h2 As Double

    If mCarregado Then Exit Sub

    For d = 1 To 7
        mTrab(d) = (d <= 5)
        mIni(d) = 6# / 24#
        mFim(d) = 22# / 24#
    Next d

    s = APS_CfgLer("CAL_SEM", CAL_SEM_PADRAO)
    itens = Split(s, "|")
    For i = 0 To UBound(itens)
        p = Split(itens(i), ";")
        If UBound(p) >= 1 Then
            d = Val(p(0))
            If d >= 1 And d <= 7 Then
                mTrab(d) = (UCase$(Left$(p(1), 1)) = "S")
                ' o horario de cada dia da semana e lido mesmo quando o dia esta sem producao: assim ele nao se perde
                ' e e o horario usado se uma DATA daquele dia da semana for liberada por excecao
                If UBound(p) >= 3 Then
                    h1 = APS_ParaHora(p(2))
                    h2 = APS_ParaHora(p(3))
                    If h1 >= 0 Then mIni(d) = h1
                    If h2 >= 0 Then mFim(d) = h2
                    If h2 = 0 Then mFim(d) = 1#
                End If
            End If
        End If
    Next i

    mNExc = 0
    ReDim mExcData(1 To 1): ReDim mExcTrab(1 To 1): ReDim mExcIni(1 To 1): ReDim mExcFim(1 To 1)
    s = APS_CfgLer("CAL_EXC", "")
    If Len(s) > 0 Then
        itens = Split(s, "|")
        ReDim mExcData(1 To UBound(itens) + 1): ReDim mExcTrab(1 To UBound(itens) + 1)
        ReDim mExcIni(1 To UBound(itens) + 1): ReDim mExcFim(1 To UBound(itens) + 1)
        For i = 0 To UBound(itens)
            p = Split(itens(i), ";")
            If UBound(p) >= 1 Then
                If APS_ParaData(p(0)) >= 0 Then
                    mNExc = mNExc + 1
                    mExcData(mNExc) = APS_ParaData(p(0))
                    mExcTrab(mNExc) = (UCase$(Left$(Trim$(p(1)), 1)) = "S")
                    ' excecao liberada sem horario proprio: usa o horario do dia da semana dessa data
                    mExcIni(mNExc) = mIni(Weekday(mExcData(mNExc), vbMonday))
                    mExcFim(mNExc) = mFim(Weekday(mExcData(mNExc), vbMonday))
                    If UBound(p) >= 3 Then
                        h1 = APS_ParaHora(p(2))
                        h2 = APS_ParaHora(p(3))
                        If h1 >= 0 Then mExcIni(mNExc) = h1
                        If h2 >= 0 Then mExcFim(mNExc) = h2
                        If h2 = 0 Then mExcFim(mNExc) = 1#
                    End If
                End If
            End If
        Next i
    End If

    mNPar = 0
    ReDim mParIni(1 To 1): ReDim mParFim(1 To 1)
    s = APS_CfgLer("CAL_PAR", "")
    If Len(s) > 0 Then
        itens = Split(s, "|")
        ReDim mParIni(1 To UBound(itens) + 1): ReDim mParFim(1 To UBound(itens) + 1)
        For i = 0 To UBound(itens)
            p = Split(itens(i), ";")
            If UBound(p) >= 1 Then
                h1 = APS_ParaDataHora(p(0))
                h2 = APS_ParaDataHora(p(1))
                If h1 >= 0 And h2 > h1 Then
                    mNPar = mNPar + 1
                    mParIni(mNPar) = h1
                    mParFim(mNPar) = h2
                End If
            End If
        Next i
    End If

    mCarregado = True
End Sub

'----------------------------------------------------------
' Intervalos de trabalho de um dia (instantes absolutos)
'----------------------------------------------------------
Public Function APS_CalIntervalos(ByVal dia As Double, ByRef ia() As Double, ByRef ib() As Double) As Long
    Dim d As Double, wd As Long, i As Long, j As Long, n As Long, m As Long
    Dim trab As Boolean, hi As Double, hf As Double
    Dim ta() As Double, tb() As Double, p As Long

    Carregar
    d = Int(dia + APS_EPS)
    wd = Weekday(d, vbMonday)
    trab = mTrab(wd): hi = mIni(wd): hf = mFim(wd)
    For i = 1 To mNExc
        If mExcData(i) = d Then trab = mExcTrab(i): hi = mExcIni(i): hf = mExcFim(i)
    Next i

    ReDim ia(1 To 20): ReDim ib(1 To 20)
    If Not trab Then Exit Function
    If hf <= hi Then hf = 1#
    n = 1
    ia(1) = d + hi
    ib(1) = d + hf

    For p = 1 To mNPar
        If mParIni(p) < d + 1# And mParFim(p) > d Then
            ReDim ta(1 To 20): ReDim tb(1 To 20)
            m = 0
            For j = 1 To n
                If mParFim(p) <= ia(j) Or mParIni(p) >= ib(j) Then
                    m = m + 1: ta(m) = ia(j): tb(m) = ib(j)
                Else
                    If mParIni(p) > ia(j) + APS_EPS Then
                        m = m + 1: ta(m) = ia(j): tb(m) = mParIni(p)
                    End If
                    If mParFim(p) < ib(j) - APS_EPS Then
                        m = m + 1: ta(m) = mParFim(p): tb(m) = ib(j)
                    End If
                End If
            Next j
            n = m
            For j = 1 To n
                ia(j) = ta(j): ib(j) = tb(j)
            Next j
        End If
    Next p
    APS_CalIntervalos = n
End Function

' Primeiro instante >= t em que ha trabalho (-1 se o calendario nao tem nenhum dia de trabalho)
Public Function APS_ProximoDisponivel(ByVal t As Double) As Double
    Dim i As Long, k As Long, n As Long, ia() As Double, ib() As Double, d0 As Double

    d0 = Int(t + APS_EPS)
    APS_ProximoDisponivel = -1
    For i = 0 To 400
        n = APS_CalIntervalos(d0 + i, ia, ib)
        For k = 1 To n
            If ib(k) > t + APS_EPS Then
                If ia(k) > t Then
                    APS_ProximoDisponivel = ia(k)
                Else
                    APS_ProximoDisponivel = t
                End If
                Exit Function
            End If
        Next k
    Next i
End Function

' Fim de uma operacao que comeca em ini (instante de trabalho) e exige dur (dias) de trabalho
Public Function APS_AdicionarTrabalho(ByVal ini As Double, ByVal dur As Double) As Double
    Dim rem_ As Double, cur As Double, i As Long, k As Long, n As Long
    Dim ia() As Double, ib() As Double, s As Double, disp As Double

    APS_AdicionarTrabalho = -1
    If dur <= APS_EPS Then APS_AdicionarTrabalho = ini: Exit Function
    rem_ = dur
    cur = ini
    For i = 0 To 1000
        n = APS_CalIntervalos(Int(ini + APS_EPS) + i, ia, ib)
        For k = 1 To n
            If ib(k) > cur + APS_EPS Then
                If ia(k) > cur Then s = ia(k) Else s = cur
                disp = ib(k) - s
                If disp >= rem_ - APS_EPS Then
                    APS_AdicionarTrabalho = APS_ArredMin(s + rem_)
                    Exit Function
                End If
                rem_ = rem_ - disp
                cur = ib(k)
            End If
        Next k
    Next i
End Function

' RECONSTRUCAO COMPATIVEL - NAO CODIGO ORIGINAL.
' Evidencia: usada por APS_FecharVaos (modAPS_Motor) para calcular o inicio de
' um Setup/Limpeza a partir do fim exigido (regra "Setup.fim = Producao.ini",
' ja existente e preservada no Stage 3). O cache de compilacao tambem lista
' este nome no bloco associado a modAPS_Calendario, reforcando que pertence
' aqui, ao lado de APS_AdicionarTrabalho. O corpo abaixo e um espelho exato
' (mesma logica, andando para tras no tempo em vez de para frente) da funcao
' APS_AdicionarTrabalho logo acima, que E codigo original existente - nao foi
' inventado nenhum criterio de calendario novo.
'
' Inicio de uma operacao que TERMINA em fim (instante de trabalho) e exigiu
' dur (dias) de trabalho.
Public Function APS_SubtrairTrabalho(ByVal fim As Double, ByVal dur As Double) As Double
    Dim rem_ As Double, cur As Double, i As Long, k As Long, n As Long
    Dim ia() As Double, ib() As Double, s As Double, disp As Double

    APS_SubtrairTrabalho = -1
    If dur <= APS_EPS Then APS_SubtrairTrabalho = fim: Exit Function
    rem_ = dur
    cur = fim
    For i = 0 To 1000
        n = APS_CalIntervalos(Int(fim + APS_EPS) - i, ia, ib)
        For k = n To 1 Step -1
            If ia(k) < cur - APS_EPS Then
                If ib(k) < cur Then s = ib(k) Else s = cur
                disp = s - ia(k)
                If disp >= rem_ - APS_EPS Then
                    APS_SubtrairTrabalho = APS_ArredMin(s - rem_)
                    Exit Function
                End If
                rem_ = rem_ - disp
                cur = ia(k)
            End If
        Next k
    Next i
End Function

' Trechos [inicio, fim] de uma operacao, apenas dentro do horario de trabalho (para desenhar os cards).
' Se a operacao estiver toda fora do calendario, devolve um unico trecho ini-fim.
Public Function APS_Segmentos(ByVal ini As Double, ByVal fim As Double, _
                              ByRef sa() As Double, ByRef sb() As Double) As Long
    Dim d As Double, n As Long, k As Long, cnt As Long, ia() As Double, ib() As Double
    Dim a As Double, b As Double, dias As Long

    ReDim sa(1 To 1): ReDim sb(1 To 1)
    dias = CLng(Int(fim - APS_EPS) - Int(ini + APS_EPS))
    If dias < 0 Or dias > 400 Then
        sa(1) = ini: sb(1) = fim: APS_Segmentos = 1: Exit Function
    End If
    ReDim sa(1 To (dias + 1) * 4 + 1): ReDim sb(1 To (dias + 1) * 4 + 1)

    For d = Int(ini + APS_EPS) To Int(fim - APS_EPS)
        n = APS_CalIntervalos(d, ia, ib)
        For k = 1 To n
            a = ia(k): b = ib(k)
            If a < ini Then a = ini
            If b > fim Then b = fim
            If b - a > APS_EPS Then
                cnt = cnt + 1
                If cnt > UBound(sa) Then Exit For
                sa(cnt) = a: sb(cnt) = b
            End If
        Next k
    Next d

    If cnt = 0 Then
        sa(1) = ini: sb(1) = fim: cnt = 1
    End If
    APS_Segmentos = cnt
End Function

Public Function APS_EmTrabalho(ByVal t As Double) As Boolean
    Dim n As Long, k As Long, ia() As Double, ib() As Double
    n = APS_CalIntervalos(Int(t + APS_EPS), ia, ib)
    For k = 1 To n
        If t >= ia(k) - APS_EPS And t < ib(k) - APS_EPS Then APS_EmTrabalho = True: Exit Function
    Next k
End Function

'----------------------------------------------------------
' Texto para o formulario de configuracao
'----------------------------------------------------------
Public Function APS_CalSemanaTexto() As String
    APS_CalSemanaTexto = APS_CfgLer("CAL_SEM", CAL_SEM_PADRAO)
End Function

Public Function APS_CalExcecoesTexto() As String
    APS_CalExcecoesTexto = Replace(APS_CfgLer("CAL_EXC", ""), "|", vbCrLf)
End Function

Public Function APS_CalParadasTexto() As String
    APS_CalParadasTexto = Replace(APS_CfgLer("CAL_PAR", ""), "|", vbCrLf)
End Function

Public Function APS_OEEPadrao() As Double
    Dim v As Double
    v = Val(Replace(APS_CfgLer("OEE_PADRAO", "100"), ",", "."))
    If v <= 0 Or v > 100 Then v = 100
    APS_OEEPadrao = v / 100#
End Function
