Attribute VB_Name = "modAPS_Motor"

Option Explicit

' mProx = horario do proximo Tick agendado (usado por APS_AgendarTick / APS_PararTick). Codigo do projeto antigo.
Private mProx As Date

' Retrato da ultima passada de APS_FecharVaosSeNecessario (so operacoes ativas): id, maquina e fim.
' Serve apenas para detectar OP excluida, encurtada ou que mudou de maquina.
Private mVaoN As Long
Private mVaoID() As String
Private mVaoMaq() As String
Private mVaoFim() As Double
Private mVaoOk As Boolean
Private mVaoIni() As Double
Private mVaoDur() As Double
Private mVaoAtr() As Double

' Intervalos de maquina LIBERADOS por um evento (OP excluida/cancelada, encurtada, movida para
' tras ou que mudou de maquina). Usados por APS_VaoAberto para separar "vao aberto por evento"
' (pode ser fechado) de "intervalo planejado" (nunca e fechado).
Private mLivN As Long
Private mLivMaq() As String
Private mLivIni() As Double
Private mLivFim() As Double
Private mLivFiltro As Boolean

' Intervalos liberados por uma movimentacao feita pelo usuario (APS_VaoPendente); entram como evento
' na proxima passada de APS_FecharVaosSeNecessario.
Private mPendN As Long
Private mPendMaq() As String
Private mPendIni() As Double
Private mPendFim() As Double

' [ADAPTACAO] True = comportamento anterior (fecha vao de TODAS as maquinas a cada atualizacao).
' False = so fecha o vao da maquina em que uma OP foi excluida, encurtada ou mudou de maquina.
Private Const APS_FECHAR_VAOS_SEMPRE As Boolean = False

'==========================================================
' APS PURAN - MOTOR
' Encaixe no calendario, cascata, relogio (Tick) e edicao manual
' em 02_Operacoes. Blocos de Tick, Encaixar, Depois e
' APS_ReencaixarLinha/APS_TrabalhoEntre: codigo REAL do projeto
' antigo (xlsm), preservado sem simplificacao.
'==========================================================


'----------------------------------------------------------
' Relogio: a cada minuto confere se algum status mudou de faixa
'----------------------------------------------------------
' Relogio: a cada minuto confere se algum status mudou de faixa
' (Planejada > Em Andamento > Concluida) ou se 02_Operacoes foi alterada/excluida
'----------------------------------------------------------
' [CORRECAO PEN-B-18] Flag de estado do ultima operacao de agendamento/cancelamento.
'   True  -> agendamento criado com sucesso ou cancelado com sucesso
'   False -> agendamento nao criado, ou cancelamento falhou
Private mProxCancelado As Boolean

Public Sub APS_AgendarTick()
    ' Cancela o agendamento anterior (mesmo horario/procedimento) antes de criar outro: evita Ticks duplicados.
    Dim cancelou As Boolean

    On Error Resume Next

    If mProx <> 0 Then
        Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!APS_Tick", , False
        If Err.Number = 0 Then
            cancelou = True
        Else
            cancelou = False
            Err.Clear
        End If
    Else
        cancelou = True    ' nada a cancelar: consideramos o estado limpo
    End If

    mProxCancelado = cancelou

    ' So cria um novo agendamento se o anterior foi cancelado com sucesso,
    ' ou se nao havia agendamento anterior. Isso evita criar um segundo evento
    ' enquanto o primeiro ainda pode estar ativo.
    If cancelou Then
        mProx = Now + TimeSerial(0, 1, 0)
        Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!APS_Tick"
        If Err.Number <> 0 Then
            mProx = 0
            mProxCancelado = False
            Err.Clear
        Else
            mProxCancelado = True
        End If
    Else
        ' o cancelamento falhou: nao criamos um novo agendamento, pois o anterior
        ' pode ainda estar ativo. Mantemos mProx inalterado para que a proxima
        ' tentativa de APS_GarantirTick tente novamente.
        Application.StatusBar = "APS: falha ao cancelar agendamento anterior do Tick."
    End If
End Sub

Public Sub APS_GarantirTick()
    If mProx = 0 Then APS_AgendarTick
End Sub

Public Sub APS_PararTick()
    On Error Resume Next
    If mProx <> 0 Then Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!APS_Tick", , False
    mProx = 0
End Sub

' O proximo Tick e SEMPRE reagendado (mesmo se a atualizacao falhar), e a falha nao e engolida:
' vai para a barra de status (sem MsgBox, para nao interromper o usuario a cada minuto).
Public Sub APS_Tick()
    Dim msg As String

    On Error GoTo Falha
    mProx = 0
    If APS_Ocupado = 0 Then
        Central_Sincronizar      ' FASE 6: mesmo relogio; no-op em modo local; falha vai para o tratador abaixo
        If APS_PrecisaAtualizar() Then
            APS_Atualizar True
            APS_UltAssin = APS_AssinaturaOps()
            APS_UltCards = APS_ContarCards()
        End If
    End If
    GoTo Fim

Falha:
    msg = Err.Description
    Err.Clear
    On Error Resume Next
    Application.StatusBar = "APS: falha no relogio (Tick): " & msg
Fim:
    On Error Resume Next
    APS_AgendarTick
End Sub

' APS_ParceiroBloco - EVIDENCIA CONFIRMADA PELO CODIGO: a propria logica de
' APS_FecharVaos, alguns paragrafos abaixo (inalterada), ja documenta a regra:
' "Len(Trim$(ops(t).idPai)) = 0" -> t e producao, parceiro e o Setup/Limpeza
' vinculado; caso contrario t e Setup/Limpeza e o parceiro e a producao cujo
' ID esta em idPai. Isso bate exatamente com o comentario do tipo tOperacao
' ("idPai: ID da producao a que a linha SETUP/LIMPEZA pertence"). Reaproveitei
' a funcao ja existente APS_AcharAuxDaProducao (modAPS_Base) para o sentido
' producao->setup/limpeza, e o mesmo idioma de busca por ID (ja usado em
' modAPS_Base/APS_ExcluirOperacao e modAPS_Cards, linha "If APS_Ig(ops(i).id,
' id) Then") para o sentido setup/limpeza->producao.
Private Function APS_ParceiroBloco(ByRef ops() As tOperacao, ByVal n As Long, ByVal t As Long) As Long
    Dim i As Long

    If Len(Trim$(ops(t).idPai)) = 0 Then
        ' t e producao: parceiro e o SETUP ou, na falta dele, a LIMPEZA vinculada
        APS_ParceiroBloco = APS_AcharAuxDaProducao(ops, n, ops(t).id, APS_TIPO_SETUP)
        If APS_ParceiroBloco = 0 Then
            APS_ParceiroBloco = APS_AcharAuxDaProducao(ops, n, ops(t).id, APS_TIPO_LIMPEZA)
        End If
    Else
        ' t e Setup/Limpeza: parceiro e a producao apontada por idPai
        For i = 1 To n
            If APS_Ig(ops(i).id, ops(t).idPai) Then
                APS_ParceiroBloco = i
                Exit Function
            End If
        Next i
    End If
End Function

' APS_Sobrepoe - RECONSTRUCAO COMPATIVEL - NAO CODIGO ORIGINAL. Publica (nao
' Private como Movivel/Depois) porque modAPS_Cards tambem a chama
' diretamente (calculo de "lane" dos cards). Evidencia: uso em
' "If APS_Sobrepoe(ops(i), ops(j)) Then" sempre comparando duas operacoes da
' MESMA maquina - teste classico de sobreposicao de intervalos [ini,fim).
Public Function APS_Sobrepoe(ByRef a As tOperacao, ByRef b As tOperacao) As Boolean
    APS_Sobrepoe = (a.ini < b.Fim - APS_EPS) And (b.ini < a.Fim - APS_EPS)
End Function

Private Function Depois(ByRef a As tOperacao, ByRef b As tOperacao) As Boolean
    If a.ini > b.ini + APS_EPS Then
        Depois = True
    ElseIf Abs(a.ini - b.ini) <= APS_EPS Then
        Depois = (StrComp(a.id, b.id, vbTextCompare) > 0)
    End If
End Function

' Movivel - EVIDENCIA FRACA (marcado explicitamente como INFERENCIA, nao fato
' confirmado). O que E confirmado: APS_Ativa ja exclui Canceladas ANTES de
' Movivel ser chamada (ops(t) so chega aqui se ja estiver ativa), entao
' Movivel tem que estar checando outra coisa. Nao existe no projeto nenhum
' campo booleano de "fixo/travado" em tOperacao, nem nenhuma trava de edicao
' por status no formulario - ou seja, NAO HA EVIDENCIA SUFICIENTE no codigo
' para determinar o criterio exato. A escolha abaixo (nao mover Em Andamento
' nem Concluida) e a interpretacao mais conservadora possivel, alinhada com a
' regra que voce mesmo reforcou repetidamente (nada em andamento ou concluido
' deve ser remanejado automaticamente) - mas e uma inferencia de seguranca,
' nao uma recuperacao de regra existente. PRECISA da sua confirmacao.
Private Function Movivel(ByRef o As tOperacao) As Boolean
    Movivel = Not (APS_Ig(o.status, APS_ST_ANDAMENTO) Or APS_Ig(o.status, APS_ST_CONCLUIDA()))
End Function

'==========================================================
Public Function APS_FecharVaos(ByRef ops() As tOperacao, ByVal n As Long, ByVal maquina As String) As Long
    Dim idx() As Long, m As Long, i As Long, j As Long, t As Long
    Dim cursor As Double, primeiro As Boolean, ini As Double, Fim As Double, dur As Double
    Dim mIni As Double, mFim As Double, limitar As Boolean
    Dim ws As Worksheet, est As Boolean
    Dim jaProcessada() As Boolean
    Dim parceiro As Long, durPar As Double, iniPar As Double, fimPar As Double
    Dim fimLogico As Double
    Dim podeAplicar As Boolean, temParceiro As Boolean
    Dim oldFimT As Double, oldFimP As Double

    limitar = APS_LimiteMes(mIni, mFim)
    ReDim jaProcessada(1 To IIf(n < 1, 1, n))

    ReDim idx(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If APS_Ativa(ops(i)) And APS_Ig(ops(i).maquina, maquina) Then
            m = m + 1: idx(m) = i
        End If
    Next i
    For i = 2 To m
        t = idx(i): j = i - 1
        Do While j >= 1
            If Depois(ops(idx(j)), ops(t)) Then
                idx(j + 1) = idx(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = t
    Next i

    primeiro = True
    For i = 1 To m
        t = idx(i)
        If jaProcessada(t) Then GoTo ProximaOperacao

        If primeiro Then
            primeiro = False
        ElseIf Movivel(ops(t)) And ops(t).ini > cursor + APS_EPS And APS_VaoAberto(maquina, cursor, ops(t).ini) Then
            ini = APS_ProximoDisponivel(cursor)
            If ini >= 0 Then
                If (Not limitar) Or ini >= mIni - APS_EPS Then
                    dur = ops(t).dur
                    If dur <= APS_EPS Then dur = ops(t).Fim - ops(t).ini
                    Fim = APS_AdicionarTrabalho(ini, dur)
                    If Fim >= 0 Then
                        If (Not limitar) Or Fim <= mFim + APS_EPS Then

                            '==================================================
                            ' BLOCO ATOMICO - valida o parceiro (Setup/Limpeza
                            ' <-> Producao) ANTES de alterar/gravar qualquer coisa.
                            '==================================================
                            parceiro = APS_ParceiroBloco(ops, n, t)
                            temParceiro = (parceiro > 0)
                            podeAplicar = True

                            If temParceiro Then
                                If jaProcessada(parceiro) Then
                                    ' o parceiro ja foi fixado nesta rodada (outro bloco) -
                                    ' nao mexe neste bloco para nao quebrar a posicao dele
                                    podeAplicar = False
                                ElseIf Not Movivel(ops(parceiro)) Then
                                    podeAplicar = False
                                Else
                                    durPar = ops(parceiro).dur
                                    If durPar <= APS_EPS Then durPar = ops(parceiro).Fim - ops(parceiro).ini

                                    If Len(Trim$(ops(t).idPai)) = 0 Then
                                        ' t = PRODUCAO, parceiro = SETUP/LIMPEZA
                                        ' regra: Setup.fim = Producao.ini
                                        fimPar = ini
                                        iniPar = APS_SubtrairTrabalho(fimPar, durPar)
                                    Else
                                        ' t = SETUP/LIMPEZA, parceiro = PRODUCAO
                                        ' regra: Producao.ini = Setup.fim
                                        iniPar = Fim
                                        fimPar = APS_AdicionarTrabalho(iniPar, durPar)
                                    End If

                                    If iniPar < 0 Or fimPar < 0 Then
                                        podeAplicar = False
                                    ElseIf limitar Then
                                        If iniPar < mIni - APS_EPS Or fimPar > mFim + APS_EPS Then podeAplicar = False
                                    End If
                                End If
                            End If

                            If podeAplicar Then
                                If ws Is Nothing Then
                                    Set ws = APS_Aba(APS_ABA_OPS)
                                    If ws Is Nothing Then Exit Function
                                    est = APS_Liberar(ws)
                                    APS_Ocupado = APS_Ocupado + 1
                                End If

                                oldFimT = ops(t).Fim
                                If temParceiro Then oldFimP = ops(parceiro).Fim

                                ops(t).ini = ini
                                ops(t).Fim = Fim

                                If temParceiro Then
                                    ops(parceiro).ini = iniPar
                                    ops(parceiro).Fim = fimPar
                                    jaProcessada(parceiro) = True
                                    APS_GravarHorario ops(parceiro)
                                End If

                                ' CASCATA PARA TRAS: o tempo que este bloco deixou livre ao ser puxado e, ele proprio,
                                ' um vao aberto por evento. Sem isto a OP seguinte via um vao "nao liberado" e a
                                ' cascata parava na primeira operacao puxada.
                                If mLivFiltro Then
                                    If oldFimT > Fim + APS_EPS Then APS_VaoLiberar maquina, Fim, oldFimT
                                    If temParceiro Then
                                        If oldFimP > fimPar + APS_EPS Then APS_VaoLiberar maquina, fimPar, oldFimP
                                    End If
                                End If

                                APS_GravarHorario ops(t)
                                jaProcessada(t) = True
                                APS_FecharVaos = APS_FecharVaos + 1
                            End If
                            ' Se podeAplicar = False: nada foi alterado. ops(t) e o
                            ' parceiro permanecem exatamente como estavam antes da
                            ' tentativa - a movimentacao parcial fica bloqueada.
                        End If
                    End If
                End If
            End If
        End If

        '======================================================
        ' ETAPA 4 - CURSOR LOGICO DO BLOCO (inalterado)
        ' O proximo encaixe considera o fim da operacao e, quando
        ' houver Setup/Limpeza, tambem o fim do parceiro.
        '======================================================
        fimLogico = ops(t).Fim
        parceiro = APS_ParceiroBloco(ops, n, t)
        If parceiro > 0 Then
            If ops(parceiro).Fim > fimLogico Then fimLogico = ops(parceiro).Fim
        End If
        If fimLogico > cursor Then cursor = fimLogico

ProximaOperacao:
    Next i

    If Not ws Is Nothing Then
        APS_Ocupado = APS_Ocupado - 1
        APS_Reproteger ws, est
    End If
End Function

Public Sub APS_FecharVaosSeNecessario(ByRef ops() As tOperacao, ByRef n As Long)
    Dim maquinas() As String, marca() As Boolean, nm As Long, i As Long, j As Long, p As Long
    Dim achou As Boolean, k As Long

    On Error GoTo Falha

    mLivN = 0
    mLivFiltro = False

    If n <= 0 Then
        mPendN = 0
        mVaoN = 0
        mVaoOk = (n = 0)
        Exit Sub
    End If

    ReDim maquinas(1 To n)
    For i = 1 To n
        If Len(Trim$(ops(i).maquina)) > 0 Then
            achou = False
            For j = 1 To nm
                If APS_Ig(maquinas(j), ops(i).maquina) Then achou = True: Exit For
            Next j
            If Not achou Then
                nm = nm + 1
                maquinas(nm) = ops(i).maquina
            End If
        End If
    Next i
    ReDim marca(1 To IIf(nm < 1, 1, nm))

    If APS_FECHAR_VAOS_SEMPRE Then
        For j = 1 To nm
            marca(j) = True
        Next j
    ElseIf mVaoOk Then
        ' EVENTO: o que ficou livre desde a ultima passada
        ReDim mLivMaq(1 To mVaoN * 2 + 1 + mPendN)
        ReDim mLivIni(1 To mVaoN * 2 + 1 + mPendN)
        ReDim mLivFim(1 To mVaoN * 2 + 1 + mPendN)
        For p = 1 To mVaoN
            achou = False
            For i = 1 To n
                If APS_Ativa(ops(i)) Then
                    If APS_Ig(ops(i).id, mVaoID(p)) Then
                        achou = True
                        If Not APS_Ig(ops(i).maquina, mVaoMaq(p)) Then
                            APS_VaoLiberar mVaoMaq(p), mVaoIni(p), mVaoFim(p)         ' mudou de maquina
                        ElseIf ops(i).Fim < mVaoFim(p) - APS_EPS Then
                            APS_VaoLiberar mVaoMaq(p), IIf(ops(i).Fim > mVaoIni(p), ops(i).Fim, mVaoIni(p)), mVaoFim(p)   ' encurtada / movida para tras
                        ElseIf Len(ops(i).tipoAux) > 0 And ops(i).ini > mVaoIni(p) + APS_EPS And Abs(ops(i).Fim - mVaoFim(p)) <= APS_EPS Then
                            ' Setup/Limpeza ENCURTADO (mesmo fim = inicio da producao, inicio mais tarde): o tempo que
                            ' sobrou antes dele e vao aberto por evento; o bloco volta para fechar o espaco
                            APS_VaoLiberar mVaoMaq(p), mVaoIni(p), ops(i).ini
                        End If
                        Exit For
                    End If
                End If
            Next i
            If Not achou Then APS_VaoLiberar mVaoMaq(p), mVaoIni(p), mVaoFim(p)       ' excluida / cancelada
        Next p
        For p = 1 To mPendN
            APS_VaoLiberar mPendMaq(p), mPendIni(p), mPendFim(p)                      ' movida pelo usuario para depois
        Next p
        APS_VaoMesclar
        For k = 1 To mLivN
            For j = 1 To nm
                If APS_Ig(maquinas(j), mLivMaq(k)) Then marca(j) = True: Exit For
            Next j
        Next k
        mLivFiltro = (mLivN > 0)
    End If

    For j = 1 To nm
        If marca(j) Then APS_FecharVaos ops, n, maquinas(j)
    Next j
    mLivFiltro = False
    mPendN = 0

    ' retrato para a proxima passada (ja com as posicoes apos o fechamento)
    ReDim mVaoID(1 To n)
    ReDim mVaoMaq(1 To n)
    ReDim mVaoIni(1 To n)
    ReDim mVaoFim(1 To n)
    ReDim mVaoDur(1 To n)
    ReDim mVaoAtr(1 To n)
    mVaoN = 0
    For i = 1 To n
        If APS_Ativa(ops(i)) Then
            mVaoN = mVaoN + 1
            mVaoID(mVaoN) = ops(i).id
            mVaoMaq(mVaoN) = ops(i).maquina
            mVaoIni(mVaoN) = ops(i).ini
            mVaoFim(mVaoN) = ops(i).Fim
            mVaoDur(mVaoN) = ops(i).dur
            mVaoAtr(mVaoN) = ops(i).atraso
        End If
    Next i
    mVaoOk = True
    Exit Sub

Falha:
    mLivFiltro = False
    mVaoOk = False
    n = -1
End Sub

' Une intervalos liberados que se tocam/sobrepoem na mesma maquina (evita contar o mesmo tempo duas vezes).
Private Sub APS_VaoMesclar()
    Dim i As Long, j As Long, mudou As Boolean

    Do
        mudou = False
        For i = 1 To mLivN - 1
            For j = i + 1 To mLivN
                If mLivN >= j And i < j Then
                    If APS_Ig(mLivMaq(i), mLivMaq(j)) Then
                        If mLivIni(j) <= mLivFim(i) + APS_EPS And mLivIni(i) <= mLivFim(j) + APS_EPS Then
                            If mLivIni(j) < mLivIni(i) Then mLivIni(i) = mLivIni(j)
                            If mLivFim(j) > mLivFim(i) Then mLivFim(i) = mLivFim(j)
                            mLivMaq(j) = mLivMaq(mLivN)
                            mLivIni(j) = mLivIni(mLivN)
                            mLivFim(j) = mLivFim(mLivN)
                            mLivN = mLivN - 1
                            mudou = True
                            Exit For
                        End If
                    End If
                End If
            Next j
            If mudou Then Exit For
        Next i
    Loop While mudou
End Sub

' Registra um intervalo de maquina liberado por um evento.
Private Sub APS_VaoLiberar(ByVal maq As String, ByVal ini As Double, ByVal Fim As Double)
    If Fim <= ini + APS_EPS Then Exit Sub
    ' os vetores crescem sob demanda (APS_FecharVaos tambem registra o que cada bloco puxado libera)
    If mLivN + 1 > UBound(mLivMaq) Then
        ReDim Preserve mLivMaq(1 To mLivN + 16)
        ReDim Preserve mLivIni(1 To mLivN + 16)
        ReDim Preserve mLivFim(1 To mLivN + 16)
    End If
    mLivN = mLivN + 1
    mLivMaq(mLivN) = maq
    mLivIni(mLivN) = ini
    mLivFim(mLivN) = Fim
End Sub

'----------------------------------------------------------
' VAO ABERTO x INTERVALO PLANEJADO
' Um vao entre duas operacoes so e fechado quando TODO o tempo de trabalho dele foi liberado
' por um evento (exclusao, cancelamento, encurtamento, movimento para tras, troca de maquina).
' Qualquer parte do vao que nao foi liberada por um evento e um intervalo planejado e nao
' e tocada. Sem filtro (chamada direta ou APS_FECHAR_VAOS_SEMPRE) = comportamento original.
'----------------------------------------------------------
Private Function APS_VaoAberto(ByVal maquina As String, ByVal a As Double, ByVal b As Double) As Boolean
    Dim total As Double, coberto As Double, k As Long, x1 As Double, x2 As Double

    If Not mLivFiltro Then APS_VaoAberto = True: Exit Function

    total = APS_TrabalhoEntre(a, b)
    If total <= APS_EPS Then APS_VaoAberto = True: Exit Function   ' sem tempo de trabalho no vao: nada planejado a proteger

    For k = 1 To mLivN
        If APS_Ig(mLivMaq(k), maquina) Then
            x1 = mLivIni(k): If x1 < a Then x1 = a
            x2 = mLivFim(k): If x2 > b Then x2 = b
            If x2 > x1 Then coberto = coberto + APS_TrabalhoEntre(x1, x2)
        End If
    Next k
    APS_VaoAberto = (coberto >= total - 2 * APS_EPS)
End Function

'==========================================================
' APS_BlocoMovivel (Private, auxiliar de APS_Encaixar) - RECONSTRUCAO
' COMPATIVEL. Usada na cascata de APS_Encaixar (adaptacao ao modelo novo): decide se um BLOCO (producao + Setup/Limpeza vinculado) pode ser
' deslocado - e um bloco so pode se mover inteiro (mesma regra de atomicidade
' ja usada em APS_FecharVaos: se o parceiro nao e movivel, nada se move).
' Reutiliza Movivel e APS_ParceiroBloco sem altera-las.
'==========================================================
Private Function APS_BlocoMovivel(ByRef ops() As tOperacao, ByVal n As Long, ByVal j As Long) As Boolean
    Dim p As Long
    If Not Movivel(ops(j)) Then Exit Function
    p = APS_ParceiroBloco(ops, n, j)
    If p > 0 Then
        If Not Movivel(ops(p)) Then Exit Function
    End If
    APS_BlocoMovivel = True
End Function

' ENCAIXE ATOMICO: guarda uma copia de ops() e so deixa o resultado se a cascata inteira foi calculada E
' validada (sem sobreposicao). Qualquer falha devolve calOk = False, ops() exatamente como entrou e mudou() zerado:
' nenhum deslocamento parcial. Quem chama so grava na planilha depois de calOk = True.
Public Function APS_Encaixar(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long, _
                             ByRef mudou() As Boolean, ByRef calOk As Boolean, _
                             Optional ByVal inicioFixo As Boolean = False, _
                             Optional ByVal ignorarMes As Boolean = False, _
                             Optional ByVal prioridade As Boolean = False) As Long
    Dim snap() As tOperacao, i As Long

    ReDim snap(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        snap(i) = ops(i)
    Next i

    APS_Encaixar = EncaixarInterno(ops, n, k, mudou, calOk, inicioFixo, ignorarMes, prioridade)

    If calOk Then
        If Not ValidarSemSobreposicao(ops, n, k, mudou) Then calOk = False
    End If

    If Not calOk Then
        For i = 1 To n
            ops(i) = snap(i)
        Next i
        ReDim mudou(1 To IIf(n < 1, 1, n))
        APS_Encaixar = 0
    End If
End Function

' Confere o resultado: nenhuma operacao deslocada (ou a escolhida) pode ficar sobreposta a outra ativa da mesma
' maquina (producao e Setup vinculados a ela nao contam entre si). Se houver, preenche APS_MotivoEncaixe.
Private Function ValidarSemSobreposicao(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long, _
                                        ByRef mudou() As Boolean) As Boolean
    Dim a As Long, b As Long

    ValidarSemSobreposicao = True
    For a = 1 To n
        If a = k Or mudou(a) Then
            If APS_Ativa(ops(a)) Then
                For b = 1 To n
                    If b <> a Then
                        If APS_Ativa(ops(b)) Then
                            If APS_Ig(ops(a).maquina, ops(b).maquina) Then
                                If APS_Sobrepoe(ops(a), ops(b)) Then
                                    If APS_ParceiroBloco(ops, n, a) <> b Then
                                        APS_MotivoEncaixe = "A movimenta" & ChrW(231) & ChrW(227) & "o foi rejeitada: " & ops(a).Produto & " " & ops(a).lote & _
                                            " (" & Format$(ops(a).ini, "dd/mm hh:nn") & " a " & Format$(ops(a).Fim, "dd/mm hh:nn") & ") ficaria sobreposta a " & _
                                            ops(b).Produto & " " & ops(b).lote & " (" & Format$(ops(b).ini, "dd/mm hh:nn") & " a " & Format$(ops(b).Fim, "dd/mm hh:nn") & _
                                            ", " & ops(b).status & ") em " & ops(a).maquina & ", que n" & ChrW(227) & "o pode ser remanejada. Nada foi alterado."
                                        ValidarSemSobreposicao = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next b
            End If
        End If
    Next a
End Function

Private Function EncaixarInterno(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long, _
                             ByRef mudou() As Boolean, ByRef calOk As Boolean, _
                             ByVal inicioFixo As Boolean, _
                             ByVal ignorarMes As Boolean, _
                             ByVal prioridade As Boolean) As Long
    Dim idx() As Long, m As Long, i As Long, j As Long, t As Long, guarda As Long
    Dim ini As Double, Fim As Double, cursor As Double, dur As Double, mover As Boolean
    Dim mesIni As Double, mesFim As Double, iniIn As Double, fimIn As Double, limitar As Boolean
    Dim pk As Long, durPar As Double, bini As Double, iniPar As Double   ' [modelo novo] Setup/Limpeza vinculado

    calOk = True
    APS_ForaDoMes = False
    APS_MotivoEncaixe = ""
    ReDim mudou(1 To IIf(n < 1, 1, n))
    If Not APS_Ativa(ops(k)) Then Exit Function

    ' limite do mes de planejamento = celula 1_Planejamento!A1 (ver APS_LimiteMes)
    iniIn = ops(k).ini
    fimIn = ops(k).Fim
    limitar = False
    If Not ignorarMes Then limitar = APS_LimiteMes(mesIni, mesFim)

    dur = ops(k).dur
    If dur <= APS_EPS Then dur = ops(k).Fim - ops(k).ini

    ' [ADAPTACAO modelo novo] Producao + Setup/Limpeza (idPai/tipoAux) formam um BLOCO:
    ' Setup.fim = Producao.ini. pk = Setup/Limpeza ativo vinculado a k (0 = sem vinculo ->
    ' o codigo abaixo se comporta exatamente como o original). Se k e ele proprio um
    ' Setup/Limpeza, e tratado como operacao comum.
    If Len(ops(k).tipoAux) = 0 Then pk = APS_ParceiroBloco(ops, n, k)
    If pk > 0 Then
        If Not APS_Ativa(ops(pk)) Then pk = 0
    End If
    If pk > 0 Then
        durPar = ops(pk).dur
        If durPar <= APS_EPS Then durPar = ops(pk).Fim - ops(pk).ini
    End If

    ini = APS_ProximoDisponivel(ops(k).ini)
    If ini < 0 Then calOk = False: Exit Function

    Do
        guarda = guarda + 1
        Fim = APS_AdicionarTrabalho(ini, dur)
        If Fim < 0 Then calOk = False: Exit Function
        ' inicio do bloco = inicio do Setup/Limpeza que precede a producao
        bini = ini
        If pk > 0 Then
            bini = APS_SubtrairTrabalho(ini, durPar)
            If bini < 0 Then bini = ini
        End If
        mover = False
        ' [modelo novo] o Setup/Limpeza tambem fica dentro do mes de planejamento: se o inicio do
        ' bloco cairia antes do dia 1, a producao avanca o tempo do Setup
        If pk > 0 And limitar Then
            If bini < mesIni - APS_EPS Then
                ini = APS_ProximoDisponivel(mesIni)
                If ini >= 0 Then ini = APS_AdicionarTrabalho(ini, durPar)
                If ini >= 0 Then ini = APS_ProximoDisponivel(ini)
                If ini < 0 Then calOk = False: Exit Function
                mover = True
                GoTo ProximaVolta
            End If
        End If
        For j = 1 To IIf(inicioFixo, 0, n)
            If j <> k And j <> pk Then
                If APS_Ativa(ops(j)) Then
                    If APS_Ig(ops(j).maquina, ops(k).maquina) Then
                        If ops(j).ini <= ini + APS_EPS Then
                            ' prioridade (posicao escolhida): so o que termina antes do inicio da producao (tempo do
                            ' Setup) empurra a producao; quem ocupa o horario escolhido e quem anda (cascata abaixo)
                            If ops(j).ini < Fim - APS_EPS And bini < ops(j).Fim - APS_EPS And _
                               ((Not prioridade) Or ops(j).Fim <= ini + APS_EPS) Then
                                ini = APS_ProximoDisponivel(ops(j).Fim)
                                If ini < 0 Then calOk = False: Exit Function
                                If pk > 0 Then
                                    ' reserva o tempo do Setup/Limpeza antes da producao
                                    ini = APS_AdicionarTrabalho(ini, durPar)
                                    If ini >= 0 Then ini = APS_ProximoDisponivel(ini)
                                    If ini < 0 Then calOk = False: Exit Function
                                End If
                                mover = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next j
ProximaVolta:
    Loop While mover And guarda < 500

    ' limite de tentativas atingido com conflito ainda aberto = FALHA (nunca aceita a posicao)
    If mover Then
        calOk = False
        APS_MotivoEncaixe = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel encontrar uma posi" & ChrW(231) & ChrW(227) & "o sem sobreposi" & ChrW(231) & ChrW(227) & "o para " & _
            ops(k).Produto & " " & ops(k).lote & " em " & ops(k).maquina & " (limite de tentativas atingido). Nada foi alterado."
        Exit Function
    End If

    ' a operacao encaixada tem que ficar dentro do mes de planejamento (nada e gravado se nao couber)
    If limitar Then
        If iniIn < mesIni - APS_EPS Or ini >= mesFim - APS_EPS Or Fim > mesFim + APS_EPS Then
            calOk = False: APS_ForaDoMes = True: APS_MesRef = mesFim - 1
            Exit Function
        End If
    End If

    ' prioridade: se o horario escolhido esta ocupado por operacao que NAO pode ser remanejada
    ' (Em Andamento/Concluida), nada e movido: informa o motivo em vez de jogar a operacao para outro lugar
    If prioridade Then
        For j = 1 To n
            If j <> k And j <> pk Then
                If APS_Ativa(ops(j)) Then
                    If APS_Ig(ops(j).maquina, ops(k).maquina) Then
                        If ops(j).ini < Fim - APS_EPS And ops(j).Fim > ini + APS_EPS Then
                            If Not APS_BlocoMovivel(ops, n, j) Then
                                APS_MotivoEncaixe = "O hor" & ChrW(225) & "rio escolhido (" & Format$(ini, "dd/mm hh:nn") & " a " & _
                                    Format$(Fim, "dd/mm hh:nn") & ") est" & ChrW(225) & " ocupado em " & ops(k).maquina & " por " & _
                                    ops(j).Produto & " " & ops(j).lote & " (" & ops(j).status & "), que n" & ChrW(227) & _
                                    "o pode ser remanejada. Escolha outro hor" & ChrW(225) & "rio. Nada foi alterado."
                                calOk = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    End If

    ops(k).ini = ini
    ops(k).Fim = Fim
    ops(k).dur = dur
    ops(k).SemDur = False

    ' [ADAPTACAO modelo novo] o Setup/Limpeza vinculado acompanha a producao (Setup.fim = Producao.ini).
    ' Quem chama grava o parceiro como qualquer outra operacao deslocada (mudou(pk) = True).
    ' Setup/Limpeza nao tem status proprio: acompanha SEMPRE a producao que foi movida (mesmo Concluida).
    If pk > 0 Then
        iniPar = APS_SubtrairTrabalho(ops(k).ini, durPar)
        If iniPar >= 0 Then
            If Abs(ops(pk).ini - iniPar) > APS_EPS Or Abs(ops(pk).Fim - ops(k).ini) > APS_EPS Then
                ops(pk).ini = iniPar
                ops(pk).Fim = ops(k).ini
                ops(pk).dur = durPar
                ops(pk).SemDur = False
                mudou(pk) = True
            End If
        End If
    End If

    ' posicao escolhida pelo usuario: quem esta no caminho avanca em cascata (blocos inteiros)
    If prioridade Then
        EncaixarInterno = CascataEscolha(ops, n, k, pk, mudou, calOk, limitar, mesFim)
        Exit Function
    End If

    ' operacoes seguintes na mesma maquina (o Setup/Limpeza de k nao entra: ja foi posicionado acima)
    ReDim idx(1 To IIf(n < 1, 1, n))
    For i = 1 To n
        If i <> k And i <> pk Then
            If APS_Ativa(ops(i)) Then
                If APS_Ig(ops(i).maquina, ops(k).maquina) Then
                    ' so entram na cascata operacoes do proprio mes; meses seguintes ficam intocados
                    If (ops(i).ini > ops(k).ini + APS_EPS Or (inicioFixo And ops(i).ini >= ops(k).ini - APS_EPS)) _
                       And (Not limitar Or ops(i).ini < mesFim - APS_EPS) Then
                        m = m + 1
                        idx(m) = i
                    End If
                End If
            End If
        End If
    Next i
    For i = 2 To m
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Depois(ops(idx(j)), ops(t)) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = t
    Next i

    cursor = ops(k).Fim
    For i = 1 To m
        j = idx(i)
        If ops(j).ini < cursor - APS_EPS Then
            If APS_BlocoMovivel(ops, n, j) Then   ' [modelo novo] bloco so se move inteiro
                dur = ops(j).dur
                If dur <= APS_EPS Then dur = ops(j).Fim - ops(j).ini
                ini = APS_ProximoDisponivel(cursor)
                If ini < 0 Then calOk = False: Exit Function
                Fim = APS_AdicionarTrabalho(ini, dur)
                If Fim < 0 Then calOk = False: Exit Function
                If limitar And Fim > mesFim + APS_EPS Then
                    ' ultrapassaria o ultimo periodo valido do mes: para, sem tocar no mes seguinte
                    calOk = False: APS_ForaDoMes = True: APS_MesRef = mesFim - 1
                    Exit Function
                End If
                ops(j).ini = ini
                ops(j).Fim = Fim
                mudou(j) = True
                EncaixarInterno = EncaixarInterno + 1
            End If
        End If
        If ops(j).Fim > cursor Then cursor = ops(j).Fim
    Next i
End Function


'----------------------------------------------------------
' CascataEscolha - usada por APS_Encaixar no modo "posicao escolhida" (prioridade).
' A operacao k ja esta posicionada onde o usuario pediu (ops(k).ini/fim). Todas as operacoes ativas da MESMA
' maquina que terminam depois do inicio de k (as que ocupam o horario escolhido e as seguintes) sao tratadas
' como BLOCOS (producao + Setup/Limpeza) e, em ordem, quem ainda estiver sobre o cursor avanca para o proximo
' tempo disponivel (calendario), reservando o Setup antes da producao (Setup.fim = Producao.ini). Quem nao
' esta no caminho nao e tocado (intervalos planejados e meses seguintes ficam como estao). Bloco que nao
' pode ser remanejado (Em Andamento/Concluida) nao e movido (o cursor passa por ele), como na cascata antiga.
' Devolve quantas operacoes foram deslocadas; calOk = False se nao coube (calendario ou mes).
'----------------------------------------------------------
Private Function CascataEscolha(ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long, ByVal pk As Long, _
                                ByRef mudou() As Boolean, ByRef calOk As Boolean, _
                                ByVal limitar As Boolean, ByVal mesFim As Double) As Long
    Dim und() As Long, par() As Long, st() As Double, visto() As Boolean
    Dim m As Long, i As Long, j As Long, u As Long, pu As Long, tu As Long, tp As Long, ts As Double
    Dim cursor As Double, ini As Double, Fim As Double, dur As Double, durS As Double, x As Double
    Dim troca As Boolean, achou As Boolean, guarda As Long, bs As Double

    ReDim und(1 To IIf(n < 1, 1, n))
    ReDim par(1 To IIf(n < 1, 1, n))
    ReDim st(1 To IIf(n < 1, 1, n))
    ReDim visto(1 To IIf(n < 1, 1, n))

    For i = 1 To n
        If i <> k And i <> pk Then
            If APS_Ativa(ops(i)) Then
                If APS_Ig(ops(i).maquina, ops(k).maquina) Then
                    If ops(i).Fim > ops(k).ini + APS_EPS Then
                        If (Not limitar) Or ops(i).ini < mesFim - APS_EPS Then
                            u = i
                            If Len(ops(i).tipoAux) > 0 Then
                                j = APS_ParceiroBloco(ops, n, i)
                                If j > 0 And j <> k Then
                                    If APS_Ativa(ops(j)) Then u = j
                                End If
                            End If
                            If Not visto(u) Then
                                visto(u) = True
                                pu = 0
                                If Len(ops(u).tipoAux) = 0 Then
                                    pu = APS_ParceiroBloco(ops, n, u)
                                    If pu > 0 Then
                                        If Not APS_Ativa(ops(pu)) Or pu = k Or pu = pk Then pu = 0
                                    End If
                                End If
                                m = m + 1
                                und(m) = u
                                par(m) = pu
                                st(m) = ops(u).ini
                                If pu > 0 Then
                                    If ops(pu).ini < st(m) Then st(m) = ops(pu).ini
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' ordem de inicio (desempate pelo ID, como Depois)
    For i = 2 To m
        tu = und(i): tp = par(i): ts = st(i)
        j = i - 1
        Do While j >= 1
            troca = False
            If st(j) > ts + APS_EPS Then
                troca = True
            ElseIf Abs(st(j) - ts) <= APS_EPS Then
                troca = (StrComp(ops(und(j)).id, ops(tu).id, vbTextCompare) > 0)
            End If
            If Not troca Then Exit Do
            und(j + 1) = und(j): par(j + 1) = par(j): st(j + 1) = st(j)
            j = j - 1
        Loop
        und(j + 1) = tu: par(j + 1) = tp: st(j + 1) = ts
    Next i

    cursor = ops(k).Fim
    For i = 1 To m
        u = und(i): pu = par(i)
        If st(i) < cursor - APS_EPS Then
            If APS_BlocoMovivel(ops, n, u) Then
                dur = ops(u).dur
                If dur <= APS_EPS Then dur = ops(u).Fim - ops(u).ini
                ini = APS_ProximoDisponivel(cursor)
                If ini < 0 Then calOk = False: Exit Function
                If pu > 0 Then
                    durS = ops(pu).dur
                    If durS <= APS_EPS Then durS = ops(pu).Fim - ops(pu).ini
                    x = APS_AdicionarTrabalho(ini, durS)
                    If x < 0 Then calOk = False: Exit Function
                    ini = APS_ProximoDisponivel(x)
                    If ini < 0 Then calOk = False: Exit Function
                End If
                Fim = APS_AdicionarTrabalho(ini, dur)
                If Fim < 0 Then calOk = False: Exit Function

                ' nao sobrepor operacao que NAO pode ser remanejada (Em Andamento/Concluida) da mesma maquina:
                ' o bloco avanca para depois dela (reservando de novo o Setup antes da producao)
                guarda = 0
                Do
                    achou = False
                    guarda = guarda + 1
                    bs = ini
                    If pu > 0 Then
                        bs = APS_SubtrairTrabalho(ini, durS)
                        If bs < 0 Then bs = ini
                    End If
                    For j = 1 To n
                        If j <> k And j <> pk And j <> u And j <> pu Then
                            If APS_Ativa(ops(j)) Then
                                If APS_Ig(ops(j).maquina, ops(k).maquina) Then
                                    If Not APS_BlocoMovivel(ops, n, j) Then
                                        If ops(j).ini < Fim - APS_EPS And ops(j).Fim > bs + APS_EPS Then
                                            ini = APS_ProximoDisponivel(ops(j).Fim)
                                            If ini < 0 Then calOk = False: Exit Function
                                            If pu > 0 Then
                                                x = APS_AdicionarTrabalho(ini, durS)
                                                If x < 0 Then calOk = False: Exit Function
                                                ini = APS_ProximoDisponivel(x)
                                                If ini < 0 Then calOk = False: Exit Function
                                            End If
                                            Fim = APS_AdicionarTrabalho(ini, dur)
                                            If Fim < 0 Then calOk = False: Exit Function
                                            achou = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next j
                Loop While achou And guarda < 200

                ' limite de tentativas com sobreposicao ainda aberta = FALHA
                If achou Then
                    calOk = False
                    APS_MotivoEncaixe = "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel reposicionar " & ops(u).Produto & " " & ops(u).lote & " sem sobrepor opera" & ChrW(231) & ChrW(227) & _
                        "o que n" & ChrW(227) & "o pode ser remanejada (limite de tentativas atingido). Nada foi alterado."
                    Exit Function
                End If

                If limitar And Fim > mesFim + APS_EPS Then
                    calOk = False: APS_ForaDoMes = True: APS_MesRef = mesFim - 1
                    Exit Function
                End If
                ops(u).ini = ini
                ops(u).Fim = Fim
                mudou(u) = True
                CascataEscolha = CascataEscolha + 1
                If pu > 0 Then
                    x = APS_SubtrairTrabalho(ini, durS)
                    If x < 0 Then x = ini - durS
                    ops(pu).ini = x
                    ops(pu).Fim = ini
                    ops(pu).dur = durS
                    ops(pu).SemDur = False
                    mudou(pu) = True
                End If
            End If
        End If
        If ops(u).Fim > cursor Then cursor = ops(u).Fim
    Next i
End Function

' Registra o espaco que uma operacao deixou ao ser movida pelo usuario para um horario posterior (sem encostar
' no antigo). Na proxima passada de APS_FecharVaosSeNecessario ele conta como vao aberto por evento.
Public Sub APS_VaoPendente(ByVal maq As String, ByVal ini As Double, ByVal Fim As Double)
    If Fim <= ini + APS_EPS Then Exit Sub
    mPendN = mPendN + 1
    ReDim Preserve mPendMaq(1 To mPendN)
    ReDim Preserve mPendIni(1 To mPendN)
    ReDim Preserve mPendFim(1 To mPendN)
    mPendMaq(mPendN) = maq
    mPendIni(mPendN) = ini
    mPendFim(mPendN) = Fim
End Sub


'----------------------------------------------------------
' Tempo de trabalho (pelo calendario existente) entre dois instantes
'----------------------------------------------------------
Public Function APS_TrabalhoEntre(ByVal ini As Double, ByVal Fim As Double) As Double
    Dim d As Double, n As Long, k As Long, ia() As Double, ib() As Double
    Dim a As Double, b As Double, tot As Double

    If Fim <= ini + APS_EPS Then Exit Function
    If Fim - ini > 400 Then Exit Function

    For d = Int(ini + APS_EPS) To Int(Fim - APS_EPS)
        n = APS_CalIntervalos(d, ia, ib)
        For k = 1 To n
            a = ia(k): b = ib(k)
            If a < ini Then a = ini
            If b > Fim Then b = Fim
            If b > a Then tot = tot + (b - a)
        Next k
    Next d
    APS_TrabalhoEntre = APS_ArredMin(tot)
End Function

'----------------------------------------------------------
' Desfaz uma edicao manual rejeitada com os valores do retrato (inicio, fim, duracao, maquina e Setup). O retrato
' e a posicao da ultima passada de APS_FecharVaosSeNecessario (= o que o APS gravou por ultimo). Devolve False se a
' operacao nao esta no retrato (ex.: linha recem-criada): nesse caso quem chama nao pode garantir o rollback.
' Chamada so de dentro de APS_OpsAlterada (aba liberada e APS_Ocupado > 0: sem eventos recursivos).
'----------------------------------------------------------
Private Function RestaurarEdicao(ByVal linha As Long, ByRef ops() As tOperacao, ByVal n As Long, ByVal k As Long) As Boolean
    Dim ws As Worksheet, vp As Long, p As Long, vs As Long
    Dim cIni As Long, cFim As Long, cDur As Long, cMaq As Long, cAtr As Long, durOr As Double

    On Error GoTo Sai
    If Not mVaoOk Then Exit Function
    For vp = 1 To mVaoN
        If APS_Ig(mVaoID(vp), ops(k).id) Then p = vp: Exit For
    Next vp
    If p = 0 Then Exit Function
    For vp = 1 To mVaoN
        If APS_Ig(mVaoID(vp), "S" & ops(k).id) Then vs = vp: Exit For
    Next vp

    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Function
    cIni = APS_Col(ws, H_INI()): cFim = APS_Col(ws, H_FIM())
    cDur = APS_Col(ws, H_DUR()): cMaq = APS_Col(ws, H_MAQ()): cAtr = APS_Col(ws, H_ATR())
    If cIni = 0 Or cFim = 0 Or linha < 2 Then Exit Function

    ws.Cells(linha, cIni).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(linha, cIni).Value2 = mVaoIni(p)
    ws.Cells(linha, cFim).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(linha, cFim).Value2 = mVaoFim(p)
    If cDur > 0 Then
        ws.Cells(linha, cDur).NumberFormat = "[h]:mm"
        durOr = mVaoDur(p)
        If mVaoAtr(p) > APS_EPS And mVaoDur(p) - mVaoAtr(p) > APS_EPS Then durOr = APS_ArredMin(mVaoDur(p) - mVaoAtr(p))
        ws.Cells(linha, cDur).Value2 = durOr           ' Duracao ORIGINAL
    End If
    If cAtr > 0 Then
        If mVaoAtr(p) > APS_EPS Then
            ws.Cells(linha, cAtr).NumberFormat = "[h]:mm"
            ws.Cells(linha, cAtr).Value2 = mVaoAtr(p)
        Else
            ws.Cells(linha, cAtr).ClearContents
        End If
    End If
    If cMaq > 0 Then ws.Cells(linha, cMaq).Value = mVaoMaq(p)
    If vs > 0 Then
        APS_GravarSetupDaOp ws, linha, mVaoDur(vs)
    Else
        APS_GravarSetupDaOp ws, linha, 0
    End If
    RestaurarEdicao = True
Sai:
End Function

'----------------------------------------------------------
' Edicao manual em 02_Operacoes -> cascata real
' bits: 1 inicio | 2 fim | 4 duracao | 8 maquina
'----------------------------------------------------------
Public Function APS_ReencaixarLinha(ByVal linha As Long, ByVal bits As Long) As Boolean
    Dim ops() As tOperacao, n As Long, k As Long, i As Long
    Dim mudou() As Boolean, calOk As Boolean, w As Double
    Dim vp As Long, ax As Long, temAnt As Boolean, antIni As Double, antFim As Double, antMaq As String
    Dim iniPedido As Double, motivo As String

    n = APS_LerOps(ops)
    If n <= 0 Then Exit Function
    For i = 1 To n
        If ops(i).linha = linha Then k = i: Exit For
    Next i
    If k = 0 Then Exit Function
    If ops(k).SemDur Then Exit Function
    If Not APS_Ativa(ops(k)) Then Exit Function

    iniPedido = ops(k).ini

    ' Posicao ANTERIOR da operacao (retrato da ultima passada de APS_FecharVaosSeNecessario): a celula ja foi editada,
    ' entao so o retrato sabe onde ela estava. Serve para tratar o espaco que ela deixa igual ao formulario.
    If mVaoOk Then
        For vp = 1 To mVaoN
            If APS_Ig(mVaoID(vp), ops(k).id) Then
                temAnt = True: antIni = mVaoIni(vp): antFim = mVaoFim(vp): antMaq = mVaoMaq(vp)
                Exit For
            End If
        Next vp
        If temAnt Then
            ax = APS_AcharAuxDaProducao(ops, n, ops(k).id, APS_TIPO_SETUP)
            If ax = 0 Then ax = APS_AcharAuxDaProducao(ops, n, ops(k).id, APS_TIPO_LIMPEZA)
            If ax > 0 Then
                For vp = 1 To mVaoN
                    If APS_Ig(mVaoID(vp), ops(ax).id) Then
                        If mVaoIni(vp) < antIni Then antIni = mVaoIni(vp)
                        Exit For
                    End If
                Next vp
            End If
        End If
    End If

    ' Setup editado na COLUNA: o inicio da PREPARACAO (bloco Setup + producao) fica onde estava e a producao passa a
    ' comecar quando a preparacao termina. Aumentar o Setup atrasa a producao; reduzir antecipa (a cascata fecha o vao).
    ' So vale se o Inicio nao foi editado junto e se ha retrato da posicao anterior.
    If (bits And 64) <> 0 And (bits And 1) = 0 And temAnt Then
        ax = APS_AcharAuxDaProducao(ops, n, ops(k).id, APS_TIPO_SETUP)
        If ax = 0 Then ax = APS_AcharAuxDaProducao(ops, n, ops(k).id, APS_TIPO_LIMPEZA)
        w = APS_ProximoDisponivel(antIni)
        If w >= 0 Then
            If ax > 0 Then
                w = APS_AdicionarTrabalho(w, ops(ax).dur)
                If w >= 0 Then w = APS_ProximoDisponivel(w)
            End If
            If w >= 0 Then
                ops(k).ini = APS_ArredMin(w)
                ops(k).Fim = APS_ArredMin(APS_AdicionarTrabalho(ops(k).ini, ops(k).dur))
                If ops(k).Fim < 0 Then ops(k).Fim = APS_ArredMin(ops(k).ini + ops(k).dur)
                bits = bits Or 1          ' posicao escolhida: a producao fica no novo inicio e as demais andam
                iniPedido = ops(k).ini
            End If
        End If
    End If

    ' Fim editado a mao: a duracao passa a ser o tempo de trabalho ate o fim informado
    If (bits And 2) <> 0 Then
        w = APS_TrabalhoEntre(ops(k).ini, ops(k).Fim)
        If w > APS_EPS Then
            ops(k).dur = w
        Else
            ops(k).dur = APS_ArredMin(ops(k).Fim - ops(k).ini)
        End If
        ' o Fim e o termino PROGRAMADO (producao + atraso): nao pode deixar menos tempo que o proprio atraso
        If ops(k).atraso > APS_EPS And ops(k).dur <= ops(k).atraso + APS_EPS Then
            motivo = "O novo Fim deixa menos tempo que o atraso registrado (" & Format$(Int(ops(k).atraso * 24# + APS_EPS), "0") & ":" & _
                     Format$(Round((ops(k).atraso * 24# - Int(ops(k).atraso * 24# + APS_EPS)) * 60#, 0), "00") & "). Reduza o atraso ou informe um Fim posterior."
            If RestaurarEdicao(linha, ops, n, k) Then
                MsgBox motivo & vbCrLf & "A altera" & ChrW(231) & ChrW(227) & "o digitada foi desfeita.", vbExclamation, APS_TITULO
            Else
                MsgBox motivo & vbCrLf & "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel restaurar o valor anterior com seguran" & ChrW(231) & "a; confira e corrija a c" & ChrW(233) & "lula.", vbExclamation, APS_TITULO
            End If
            Exit Function
        End If
    End If

    ' inicio ou maquina digitados a mao = posicao escolhida: a operacao fica e as demais andam em cascata
    APS_Encaixar ops, n, k, mudou, calOk, False, False, ((bits And 9) <> 0)
    If Not calOk Then
        ' Falha: nada foi gravado pelo APS (o encaixe e atomico). A celula ja foi alterada pelo usuario; ela e
        ' restaurada com os valores do retrato da ultima passada. Application.Undo so e usado se nao houver retrato,
        ' porque o proprio VBA (protecao da aba, formatacao, IDs) invalida a pilha de desfazer.
        motivo = ""
        If Len(APS_MotivoEncaixe) > 0 Then
            motivo = APS_MotivoEncaixe
        ElseIf APS_ForaDoMes Then
            motivo = APS_MsgForaDoMes()
        Else
            motivo = "O calend" & ChrW(225) & "rio n" & ChrW(227) & "o possui nenhum per" & ChrW(237) & "odo de trabalho. Ajuste em CALEND" & ChrW(193) & "RIO."
        End If
        If Len(motivo) > 0 Then
            If RestaurarEdicao(linha, ops, n, k) Then
                MsgBox motivo & vbCrLf & "A altera" & ChrW(231) & ChrW(227) & "o digitada foi desfeita.", vbExclamation, APS_TITULO
            Else
                On Error Resume Next
                Application.Undo
                Err.Clear
                On Error GoTo 0
                MsgBox motivo & vbCrLf & "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel restaurar o valor anterior com seguran" & ChrW(231) & "a; confira e corrija a c" & ChrW(233) & "lula.", vbExclamation, APS_TITULO
            End If
        End If
        Exit Function
    End If

    APS_GravarOp ops(k)
    For i = 1 To n
        If i <> k Then
            If mudou(i) Then APS_GravarHorario ops(i)
        End If
    Next i

    ' operacao movida a mao para um horario POSTERIOR sem encostar no antigo: o espaco que ela deixou e vao aberto
    ' por evento e pode ser fechado (mesma regra do formulario)
    If temAnt And ((bits And 9) <> 0) Then
        If APS_Ig(antMaq, ops(k).maquina) And ops(k).ini > antFim + APS_EPS Then APS_VaoPendente antMaq, antIni, antFim
    End If

    ' nada de ajuste silencioso: se o inicio digitado nao pode ser atendido, informa o motivo
    If Abs(ops(k).ini - iniPedido) > APS_EPS Then
        MsgBox "In" & ChrW(237) & "cio ajustado: " & Format$(iniPedido, "dd/mm hh:nn") & " > " & Format$(ops(k).ini, "dd/mm hh:nn") & vbCrLf & _
               "O hor" & ChrW(225) & "rio digitado n" & ChrW(227) & "o estava dispon" & ChrW(237) & "vel (fora do expediente, m" & ChrW(225) & "quina ocupada por opera" & ChrW(231) & ChrW(227) & _
               "o que n" & ChrW(227) & "o pode ser remanejada, ou tempo de Setup/Limpeza a reservar antes da produ" & ChrW(231) & ChrW(227) & "o).", vbInformation, APS_TITULO
    End If
    APS_ReencaixarLinha = True
End Function


