Attribute VB_Name = "modAPS_Central"
Option Explicit
'==========================================================
' APS PURAN - CENTRAL (integracao SharePoint Online) - FASE 6
'
' ESTADO HONESTO DESTE MODULO
'  - Partes PURAS (tempo, JSON, mapeamento, chave, sobreposicao, regras de conflito, fila):
'    espelhadas em Python (testes_locais/central_espelho.py) e cobertas pela macro Central_AutoTeste.
'    O codigo VBA NAO foi compilado nem executado aqui (nao ha Excel). Compile e rode o autoteste.
'  - Transporte HTTP e AUTENTICACAO: escritos, NUNCA executados contra SharePoint/Graph.
'    Endpoints, cabecalhos If-Match/ETag e filtro por Modified sao PREMISSAS a provar no tenant de teste.
'  - Modo padrao = LOCAL: com ele este modulo nao faz nada e o APS se comporta como antes.
'  - Nenhum segredo, senha ou token e gravado na planilha, em nomes ou em log. Token so em memoria.
'  - A sincronizacao NUNCA altera a estrutura da planilha (colunas/abas): so escreve valores em
'    colunas existentes. Abas auxiliares (APS_FILA) so sao criadas por Central_Instalar (acao do admin).
'
' REGRAS DE NEGOCIO (decididas pelo usuario)
'  - Fuso: America/Sao_Paulo = UTC-3 fixo. No SharePoint as datas ficam em UTC (Z).
'  - Conflito 412 em alteracao AUTOMATICA de status: descarta a alteracao, sincroniza e recalcula.
'  - Operacoes sobrepostas na mesma maquina: status existente "Conflito"; NAO move, NAO exclui,
'    NAO sobrescreve; registra em auditoria; o planejador decide.
'  - Intervalo: padrao 60 s, configuravel. O relogio do Motor (APS_Tick) e fixo em 1 minuto e
'    PROTEGIDO (nao alterado); portanto o intervalo efetivo e multiplo de 60 s, minimo 60.
'==========================================================

Public Const CEN_VERSAO As String = "6.0-local"
Public Const CEN_ABA_FILA As String = "APS_FILA"
Public Const CEN_MODO_LOCAL As String = "LOCAL"
Public Const CEN_MODO_CONECTADO As String = "CONECTADO"
Public Const CEN_OFFSET_HORAS As Long = 3          ' America/Sao_Paulo (UTC-3) -> UTC = local + 3 h
Public Const CEN_LOTE_IDS As Long = 20             ' tamanho do bloco de ApsId reservado do contador central
Public Const CEN_SCHEMA_ESPERADO As String = "1"

Private mModo As String
Private mModoLido As Boolean
Private mAplicando As Boolean                      ' True enquanto o pull escreve na planilha (desliga marcacao de sujo)
Private mEmSync As Boolean
Private mUltSync As Double
Private mCursor As String                          ' ultimo "Modificado" (UTC ISO) aprendido num PULL
Private mUltErro As String
Private mAvisouSemId As Boolean

'==========================================================
' CONFIGURACAO / ESTADO
'==========================================================
Public Sub Central_RecarregarConfig()
    mModoLido = False
End Sub

Public Function Central_Modo() As String
    If Not mModoLido Then
        mModo = UCase$(Trim$(APS_CfgLer("CEN_MODO", CEN_MODO_LOCAL)))
        If mModo <> CEN_MODO_CONECTADO Then mModo = CEN_MODO_LOCAL
        mModoLido = True
    End If
    Central_Modo = mModo
End Function

' True quando a planilha esta em modo conectado (independe de o pull estar escrevendo)
Public Function Central_Conectado() As Boolean
    Central_Conectado = (Central_Modo() = CEN_MODO_CONECTADO)
End Function

' True quando os gravadores devem marcar a linha como "suja" (conectado e nao aplicando pull)
Public Function Central_Ativo() As Boolean
    If mAplicando Then Exit Function
    Central_Ativo = Central_Conectado()
End Function

' Em modo conectado, rotinas que inserem/movem colunas ou migram dados ficam DESLIGADAS
Public Function Central_EstruturaTravada() As Boolean
    Central_EstruturaTravada = Central_Conectado()
End Function

Public Sub Central_DefinirModo(ByVal modo As String)
    modo = UCase$(Trim$(modo))
    If modo <> CEN_MODO_CONECTADO Then modo = CEN_MODO_LOCAL
    APS_CfgGravar "CEN_MODO", modo
    Central_RecarregarConfig
End Sub

Public Function Central_IntervaloSeg() As Long
    Dim v As Long
    v = CLng(Val(APS_CfgLer("CEN_INT", "60")))
    If v < 60 Then v = 60
    If v > 3600 Then v = 3600
    Central_IntervaloSeg = v
End Function

Public Sub Central_DefinirIntervalo(ByVal seg As Long)
    If seg < 60 Then seg = 60
    If seg > 3600 Then seg = 3600
    APS_CfgGravar "CEN_INT", CStr(seg)
End Sub

Public Function Central_UltimoErro() As String
    Central_UltimoErro = mUltErro
End Function

' Texto curto para a barra de status
Public Function Central_Estado() As String
    Dim p As Long
    If Not Central_Conectado() Then
        Central_Estado = "APS: modo local"
        Exit Function
    End If
    p = Central_FilaPendentes()
    Central_Estado = "APS conectado | pendentes: " & p
    If Len(mUltErro) > 0 Then Central_Estado = Central_Estado & " | " & Left$(mUltErro, 80)
End Function

'==========================================================
' TEMPO  (serial do Excel, hora de Sao Paulo  <->  ISO 8601 UTC)
'==========================================================
Private Function Z2(ByVal v As Long) As String
    Z2 = Right$("00" & CStr(v), 2)
End Function

' Arredonda o serial ao segundo e devolve "yyyy-mm-ddThh:mm:ssZ" em UTC
Public Function Central_SerialParaIso(ByVal serial As Double) As String
    Dim sec As Double, dia As Double, resto As Double, d As Date
    Dim hh As Long, mi As Long, ss As Long

    If serial <= 0 Then Exit Function
    sec = Int(serial * 86400# + 0.5) + CEN_OFFSET_HORAS * 3600#
    dia = Int(sec / 86400#)
    resto = sec - dia * 86400#
    d = CDate(dia)
    hh = CLng(Int(resto / 3600#))
    mi = CLng(Int((resto - hh * 3600#) / 60#))
    ss = CLng(resto - hh * 3600# - mi * 60#)
    Central_SerialParaIso = Right$("0000" & CStr(Year(d)), 4) & "-" & Z2(Month(d)) & "-" & Z2(Day(d)) & _
                            "T" & Z2(hh) & ":" & Z2(mi) & ":" & Z2(ss) & "Z"
End Function

' "yyyy-mm-ddThh:mm:ssZ" (UTC) -> serial local. Devolve 0 se o texto nao for valido.
Public Function Central_IsoParaSerial(ByVal iso As String) As Double
    Dim a As Long, m As Long, d As Long, h As Long, n As Long, s As Long, base As Double

    iso = Trim$(iso)
    If Len(iso) < 19 Then Exit Function
    If Mid$(iso, 5, 1) <> "-" Or Mid$(iso, 8, 1) <> "-" Or Mid$(iso, 11, 1) <> "T" _
       Or Mid$(iso, 14, 1) <> ":" Or Mid$(iso, 17, 1) <> ":" Then Exit Function
    If Not IsNumeric(Mid$(iso, 1, 4)) Or Not IsNumeric(Mid$(iso, 6, 2)) Or Not IsNumeric(Mid$(iso, 9, 2)) _
       Or Not IsNumeric(Mid$(iso, 12, 2)) Or Not IsNumeric(Mid$(iso, 15, 2)) Or Not IsNumeric(Mid$(iso, 18, 2)) Then Exit Function

    a = CLng(Mid$(iso, 1, 4)): m = CLng(Mid$(iso, 6, 2)): d = CLng(Mid$(iso, 9, 2))
    h = CLng(Mid$(iso, 12, 2)): n = CLng(Mid$(iso, 15, 2)): s = CLng(Mid$(iso, 18, 2))
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or h > 23 Or n > 59 Or s > 59 Then Exit Function

    base = CDbl(DateSerial(a, m, d))
    base = base * 86400# + h * 3600# + n * 60# + s - CEN_OFFSET_HORAS * 3600#
    Central_IsoParaSerial = base / 86400#
End Function

' duracao em dias (serial) -> segundos inteiros (evita ponto flutuante no SharePoint)
Public Function Central_DiasParaSeg(ByVal dias As Double) As Double
    If dias <= 0 Then
        Central_DiasParaSeg = 0
    Else
        Central_DiasParaSeg = Int(dias * 86400# + 0.5)
    End If
End Function

Public Function Central_SegParaDias(ByVal seg As Double) As Double
    Central_SegParaDias = seg / 86400#
End Function

'==========================================================
' JSON MINIMO (objeto plano, strings/numeros/booleanos/null, objetos e arrays como texto bruto)
'==========================================================
Public Function Central_JsonEsc(ByVal s As String) As String
    Dim i As Long, c As Long, r As String, ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34: r = r & "\" & Chr$(34)
            Case 92: r = r & "\\"
            Case 8: r = r & "\b"
            Case 9: r = r & "\t"
            Case 10: r = r & "\n"
            Case 12: r = r & "\f"
            Case 13: r = r & "\r"
            Case Is < 32, Is > 126: r = r & "\u" & Right$("0000" & LCase$(Hex$(c)), 4)
            Case Else: r = r & ch
        End Select
    Next i
    Central_JsonEsc = r
End Function

Private Sub J_Esp(ByRef js As String, ByRef p As Long)
    Do While p <= Len(js)
        Select Case Mid$(js, p, 1)
            Case " ", vbTab, vbCr, vbLf: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

' le uma string JSON a partir de p (que aponta para a aspa de abertura); devolve o texto sem escapes
Private Function J_Str(ByRef js As String, ByRef p As Long) As String
    Dim r As String, ch As String, e As String

    p = p + 1
    Do While p <= Len(js)
        ch = Mid$(js, p, 1)
        If ch = Chr$(34) Then
            p = p + 1
            J_Str = r
            Exit Function
        ElseIf ch = "\" Then
            e = Mid$(js, p + 1, 1)
            Select Case e
                Case "n": r = r & vbLf
                Case "r": r = r & vbCr
                Case "t": r = r & vbTab
                Case "b": r = r & Chr$(8)
                Case "f": r = r & Chr$(12)
                Case "u"
                    r = r & ChrW(CLng("&H" & Mid$(js, p + 2, 4)))
                    p = p + 4
                Case Else: r = r & e
            End Select
            p = p + 2
        Else
            r = r & ch
            p = p + 1
        End If
    Loop
    J_Str = r
End Function

' le um valor a partir de p. tipo: "s" string, "r" bruto (numero, true/false/null, objeto, array)
Private Function J_Val(ByRef js As String, ByRef p As Long, ByRef tipo As String) As String
    Dim ch As String, ini As Long, prof As Long, emStr As Boolean, c As String

    J_Esp js, p
    ch = Mid$(js, p, 1)
    If ch = Chr$(34) Then
        tipo = "s"
        J_Val = J_Str(js, p)
        Exit Function
    End If

    tipo = "r"
    ini = p
    If ch = "{" Or ch = "[" Then
        Do While p <= Len(js)
            c = Mid$(js, p, 1)
            If emStr Then
                If c = "\" Then
                    p = p + 1
                ElseIf c = Chr$(34) Then
                    emStr = False
                End If
            Else
                If c = Chr$(34) Then
                    emStr = True
                ElseIf c = "{" Or c = "[" Then
                    prof = prof + 1
                ElseIf c = "}" Or c = "]" Then
                    prof = prof - 1
                    If prof = 0 Then
                        p = p + 1
                        Exit Do
                    End If
                End If
            End If
            p = p + 1
        Loop
    Else
        Do While p <= Len(js)
            c = Mid$(js, p, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Or c = vbCr Or c = vbLf Or c = vbTab Then Exit Do
            p = p + 1
        Loop
    End If
    J_Val = Mid$(js, ini, p - ini)
End Function

' valor da chave de 1o nivel de um objeto. achou=False se nao existir. null vira "".
Public Function Central_JsonGet(ByVal js As String, ByVal chave As String, Optional ByRef achou As Boolean) As String
    Dim p As Long, k As String, v As String, tipo As String

    achou = False
    p = InStr(js, "{")
    If p = 0 Then Exit Function
    p = p + 1
    Do
        J_Esp js, p
        If p > Len(js) Then Exit Function
        If Mid$(js, p, 1) = "}" Then Exit Function
        If Mid$(js, p, 1) <> Chr$(34) Then Exit Function
        k = J_Str(js, p)
        J_Esp js, p
        If Mid$(js, p, 1) <> ":" Then Exit Function
        p = p + 1
        v = J_Val(js, p, tipo)
        If k = chave Then
            achou = True
            If tipo = "r" And v = "null" Then v = ""
            Central_JsonGet = v
            Exit Function
        End If
        J_Esp js, p
        If Mid$(js, p, 1) = "," Then p = p + 1
    Loop
End Function

' separa um array JSON bruto "[{...},{...}]" em elementos brutos. Devolve a quantidade.
Public Function Central_JsonItens(ByVal arr As String, ByRef itens() As String) As Long
    Dim p As Long, n As Long, v As String, tipo As String

    ReDim itens(1 To 1)
    p = InStr(arr, "[")
    If p = 0 Then Exit Function
    p = p + 1
    Do
        J_Esp arr, p
        If p > Len(arr) Then Exit Do
        If Mid$(arr, p, 1) = "]" Then Exit Do
        v = J_Val(arr, p, tipo)
        n = n + 1
        If n > UBound(itens) Then ReDim Preserve itens(1 To n * 2)
        itens(n) = v
        J_Esp arr, p
        If Mid$(arr, p, 1) = "," Then p = p + 1
    Loop
    Central_JsonItens = n
End Function

Private Function J_Num(ByVal v As Double) As String
    J_Num = Replace(CStr(v), ",", ".")
End Function

Private Function J_Campo(ByVal nome As String, ByVal valor As String, ByVal texto As Boolean) As String
    If texto Then
        J_Campo = Chr$(34) & nome & Chr$(34) & ":" & Chr$(34) & Central_JsonEsc(valor) & Chr$(34)
    Else
        J_Campo = Chr$(34) & nome & Chr$(34) & ":" & valor
    End If
End Function

'==========================================================
' MAPEAMENTO  tOperacao  <->  campos da lista APS_Operacoes
'==========================================================
' Apenas operacoes de PRODUCAO (tipoAux vazio). O Setup viaja no campo SetupSeg da propria producao.
Public Function Central_OpParaCampos(ByRef o As tOperacao, ByVal chave As String, ByVal excluido As Boolean) As String
    Dim r As String

    r = "{" & J_Campo("Title", o.Produto & " | " & o.lote, True)
    r = r & "," & J_Campo("ApsId", CStr(CLng(Val(o.id))), False)
    r = r & "," & J_Campo("ChaveCliente", chave, True)
    r = r & "," & J_Campo("CodigoProduto", o.codigo, True)
    r = r & "," & J_Campo("OP", o.ordem, True)
    r = r & "," & J_Campo("Lote", o.lote, True)
    r = r & "," & J_Campo("Maquina", o.maquina, True)
    If o.ini > 0 Then r = r & "," & J_Campo("Inicio", Central_SerialParaIso(o.ini), True)
    If o.Fim > 0 And Not o.SemDur Then r = r & "," & J_Campo("Fim", Central_SerialParaIso(o.Fim), True)
    r = r & "," & J_Campo("Status", o.status, True)
    r = r & "," & J_Campo("QtdCaixas", J_Num(o.caixas), False)
    r = r & "," & J_Campo("QtdComprimidos", J_Num(o.comp), False)
    r = r & "," & J_Campo("DuracaoSeg", J_Num(Central_DiasParaSeg(APS_DurOriginal(o))), False)
    r = r & "," & J_Campo("AtrasoSeg", J_Num(Central_DiasParaSeg(o.atraso)), False)
    r = r & "," & J_Campo("OEE", J_Num(o.oee), False)
    r = r & "," & J_Campo("VelBase", J_Num(o.velBase), False)
    r = r & "," & J_Campo("Observacoes", o.obs, True)
    If excluido Then
        r = r & "," & J_Campo("Excluido", "true", False)
    Else
        r = r & "," & J_Campo("Excluido", "false", False)
    End If
    Central_OpParaCampos = r & "}"
End Function

' Campos -> tOperacao. Devolve "" se valido, ou o motivo da rejeicao (item NAO deve ser aplicado).
Public Function Central_CamposParaOp(ByVal campos As String, ByRef o As tOperacao, ByRef setupSeg As Double, _
                                     ByRef excluido As Boolean) As String
    Dim v As String, a As Boolean, ini As Double, fim As Double, ok As Boolean, lst As Variant, i As Long

    o.linha = 0
    o.id = ""
    o.tipoAux = "": o.idPai = ""

    v = Central_JsonGet(campos, "ApsId", a)
    If Not a Or Not IsNumeric(v) Then Central_CamposParaOp = "ApsId ausente/invalido": Exit Function
    o.id = CStr(CLng(Val(v)))
    If CLng(Val(v)) <= 0 Then Central_CamposParaOp = "ApsId <= 0": Exit Function

    excluido = (LCase$(Central_JsonGet(campos, "Excluido")) = "true")

    o.codigo = Central_JsonGet(campos, "CodigoProduto")
    o.ordem = Central_JsonGet(campos, "OP")
    o.lote = Central_JsonGet(campos, "Lote")
    o.maquina = Central_JsonGet(campos, "Maquina")
    o.obs = Central_JsonGet(campos, "Observacoes")
    o.Produto = Central_JsonGet(campos, "Title")
    If InStr(o.Produto, " | ") > 0 Then o.Produto = Left$(o.Produto, InStrRev(o.Produto, " | ") - 1)

    ini = Central_IsoParaSerial(Central_JsonGet(campos, "Inicio"))
    fim = Central_IsoParaSerial(Central_JsonGet(campos, "Fim"))
    o.ini = ini
    o.Fim = fim
    o.SemDur = (fim <= 0)

    o.status = Central_JsonGet(campos, "Status")
    o.caixas = Val(Central_JsonGet(campos, "QtdCaixas"))
    o.comp = Val(Central_JsonGet(campos, "QtdComprimidos"))
    o.oee = Val(Central_JsonGet(campos, "OEE"))
    o.velBase = Val(Central_JsonGet(campos, "VelBase"))
    o.atraso = Central_SegParaDias(Val(Central_JsonGet(campos, "AtrasoSeg")))
    o.dur = Central_SegParaDias(Val(Central_JsonGet(campos, "DuracaoSeg"))) + o.atraso
    setupSeg = Val(Central_JsonGet(campos, "SetupSeg"))

    If excluido Then Exit Function          ' item excluido: so o ApsId importa

    ' validacao (item invalido NAO e aplicado)
    If Len(Trim$(o.maquina)) = 0 Then Central_CamposParaOp = "maquina vazia": Exit Function
    If ini <= 0 Then Central_CamposParaOp = "Inicio invalido": Exit Function
    If fim > 0 And fim < ini Then Central_CamposParaOp = "Fim anterior ao Inicio": Exit Function
    lst = APS_ListaStatus()
    For i = LBound(lst) To UBound(lst)
        If StrComp(CStr(lst(i)), o.status, vbTextCompare) = 0 Then ok = True: o.status = CStr(lst(i))
    Next i
    If Not ok Then Central_CamposParaOp = "status invalido: " & o.status: Exit Function
End Function

' Chave unica de idempotencia (GUID quando possivel; senao data/hora + aleatorio)
Public Function Central_NovaChave() As String
    Dim g As String

    On Error Resume Next
    g = CreateObject("Scriptlet.TypeLib").GUID
    On Error GoTo 0
    If Len(g) < 30 Then
        Randomize
        g = Format$(Now, "yyyymmddhhnnss") & Hex$(CLng(Rnd * 2147483647#)) & Hex$(CLng(Timer * 1000#))
    End If
    g = Replace(Replace(Replace(g, "{", ""), "}", ""), vbNullChar, "")
    Central_NovaChave = "APS-" & g
End Function

'==========================================================
' SOBREPOSICAO NA MESMA MAQUINA  (regra: sinalizar "Conflito", nunca mover/excluir/sobrescrever)
'==========================================================
' Intervalos [ini,fim) em serial. Encostar (fimA = iniB) NAO e sobreposicao.
Public Function Central_Sobrepoe(ByVal iniA As Double, ByVal fimA As Double, ByVal iniB As Double, ByVal fimB As Double) As Boolean
    If fimA <= iniA Or fimB <= iniB Then Exit Function
    Central_Sobrepoe = (iniA < fimB - APS_EPS) And (iniB < fimA - APS_EPS)
End Function

' Marca "Conflito" nas producoes que se sobrepoem na mesma maquina. NAO altera datas, NAO exclui.
' Devolve quantas operacoes foram marcadas. Nao marca operacoes Concluidas/Canceladas.
' So atua em modo conectado (em modo local o APS continua como antes).
Public Function Central_SinalizarSobreposicoes(ByRef ops() As tOperacao, ByVal n As Long) As Long
    Dim i As Long, j As Long, marc() As Boolean, c As Long

    If n < 2 Then Exit Function
    ReDim marc(1 To n)
    For i = 1 To n - 1
        If Len(ops(i).tipoAux) = 0 And Central_Participa(ops(i)) Then
            For j = i + 1 To n
                If Len(ops(j).tipoAux) = 0 And Central_Participa(ops(j)) Then
                    If StrComp(Trim$(ops(i).maquina), Trim$(ops(j).maquina), vbTextCompare) = 0 Then
                        If Central_Sobrepoe(ops(i).ini, ops(i).Fim, ops(j).ini, ops(j).Fim) Then
                            marc(i) = True: marc(j) = True
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    For i = 1 To n
        If marc(i) Then
            If StrComp(ops(i).status, APS_ST_CONFLITO, vbTextCompare) <> 0 Then
                ops(i).status = APS_ST_CONFLITO
                c = c + 1
            End If
        End If
    Next i
    Central_SinalizarSobreposicoes = c
End Function

Private Function Central_Participa(ByRef o As tOperacao) As Boolean
    If o.SemDur Then Exit Function
    If o.Fim <= o.ini Then Exit Function
    If StrComp(o.status, APS_ST_CANCELADA, vbTextCompare) = 0 Then Exit Function
    If StrComp(o.status, APS_ST_CONCLUIDA(), vbTextCompare) = 0 Then Exit Function
    Central_Participa = True
End Function

' Decisao ao receber 412 (ETag obsoleto) no envio. Devolve:
'  "DESCARTAR" = alteracao automatica de status: descarta, sincroniza, recalcula (decisao 3)
'  "MANTER"    = alteracao do usuario: NAO sobrescreve; fica pendente com aviso para o responsavel decidir
Public Function Central_DecisaoConflitoEtag(ByVal acao As String) As String
    If UCase$(acao) = "STATUSAUTO" Then
        Central_DecisaoConflitoEtag = "DESCARTAR"
    Else
        Central_DecisaoConflitoEtag = "MANTER"
    End If
End Function

'==========================================================
' FILA LOCAL (aba oculta APS_FILA) + MAPA ApsId -> ItemId/ETag
'   Fila:  A ApsId | B Acao (UPSERT/EXCLUIR/STATUSAUTO) | C Chave | D Tentativas | E Momento | F Erro
'   Mapa:  H ApsId | I ItemId | J ETag | K ModificadoUtc
'==========================================================
Public Function Central_AbaFila() As Worksheet
    Set Central_AbaFila = APS_Aba(CEN_ABA_FILA)
End Function

' Acao do administrador: cria a aba oculta (a sincronizacao NUNCA cria abas)
Public Sub Central_Instalar()
    Dim ws As Worksheet

    Set ws = Central_AbaFila()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CEN_ABA_FILA
        ws.Range("A1:F1").Value = Array("ApsId", "Acao", "Chave", "Tentativas", "Momento", "Erro")
        ws.Range("H1:K1").Value = Array("ApsId", "ItemId", "ETag", "ModificadoUtc")
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

Private Function Central_FilaLinha(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim u As Long, r As Long
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To u
        If CStr(ws.Cells(r, 1).Value) = id Then Central_FilaLinha = r: Exit Function
    Next r
End Function

' Marca (ou atualiza) um item pendente. EXCLUIR prevalece; UPSERT prevalece sobre STATUSAUTO.
Public Function Central_FilaMarcar(ByVal id As String, ByVal acao As String) As Boolean
    Dim ws As Worksheet, r As Long, ant As String, est As Boolean

    If Len(id) = 0 Then Exit Function
    Set ws = Central_AbaFila()
    If ws Is Nothing Then Exit Function
    On Error GoTo Falha
    est = APS_Liberar(ws)
    r = Central_FilaLinha(ws, id)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = id
        ws.Cells(r, 2).Value = acao
        ws.Cells(r, 3).Value = Central_NovaChave()
        ws.Cells(r, 4).Value = 0
    Else
        ant = CStr(ws.Cells(r, 2).Value)
        If ant = "EXCLUIR" Then
            ' continua EXCLUIR
        ElseIf acao = "EXCLUIR" Then
            ws.Cells(r, 2).Value = "EXCLUIR"
        ElseIf ant = "STATUSAUTO" And acao = "UPSERT" Then
            ws.Cells(r, 2).Value = "UPSERT"
        End If
    End If
    ws.Cells(r, 5).Value = Now
    Central_FilaMarcar = True
Saida:
    On Error Resume Next
    APS_Reproteger ws, est
    Exit Function
Falha:
    mUltErro = "fila: " & Err.Description
    Resume Saida
End Function

Public Function Central_FilaPendentes() As Long
    Dim ws As Worksheet
    Set ws = Central_AbaFila()
    If ws Is Nothing Then Exit Function
    Central_FilaPendentes = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    If Central_FilaPendentes < 0 Then Central_FilaPendentes = 0
End Function

Private Sub Central_FilaRemover(ByVal ws As Worksheet, ByVal r As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).ClearContents
    ' compacta: move a ultima linha preenchida para o buraco
    Dim u As Long
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u > r Then
        ws.Range(ws.Cells(u, 1), ws.Cells(u, 6)).Copy ws.Cells(r, 1)
        ws.Range(ws.Cells(u, 1), ws.Cells(u, 6)).ClearContents
    End If
    Application.CutCopyMode = False
End Sub

' Marca a linha de 02_Operacoes como pendente (chamado pelos gravadores do Base)
Public Sub Central_MarcarLinha(ByVal linha As Long, Optional ByVal acao As String = "UPSERT")
    Dim ws As Worksheet, cId As Long, id As String

    If Not Central_Ativo() Then Exit Sub
    If linha < 2 Then Exit Sub
    Set ws = APS_Aba(APS_ABA_OPS)
    If ws Is Nothing Then Exit Sub
    cId = APS_Col(ws, H_ID())
    If cId = 0 Then Exit Sub
    id = CStr(ws.Cells(linha, cId).Value)
    If Len(id) = 0 Then Exit Sub
    Central_FilaMarcar id, acao
End Sub

Public Sub Central_MarcarId(ByVal id As String, Optional ByVal acao As String = "UPSERT")
    If Not Central_Ativo() Then Exit Sub
    Central_FilaMarcar id, acao
End Sub

'==========================================================
' ApsId: contador central com bloco reservado (OCC). Offline usa o bloco local; esgotado -> "" (nao cria)
'==========================================================
Public Function Central_PodeCriar() As Boolean
    Dim s As String
    If Not Central_Conectado() Then Central_PodeCriar = True: Exit Function
    s = Central_ProximoIdCentral(False)
    Central_PodeCriar = (Len(s) > 0)
    If Not Central_PodeCriar And Not mAvisouSemId Then
        mAvisouSemId = True
        MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel reservar um n" & ChrW(250) & "mero de opera" & ChrW(231) & ChrW(227) & "o na lista central (sem conex" & ChrW(227) & "o?)." & vbCrLf & _
               "A cria" & ChrW(231) & ChrW(227) & "o de novas opera" & ChrW(231) & ChrW(245) & "es fica indispon" & ChrW(237) & "vel at" & ChrW(233) & " a conex" & ChrW(227) & "o voltar. Edi" & ChrW(231) & ChrW(245) & "es continuam na fila.", _
               vbExclamation, APS_TITULO
    End If
End Function

' consome=True entrega o proximo ApsId; False apenas verifica se ha como entregar
Public Function Central_ProximoIdCentral(ByVal consome As Boolean) As String
    Dim at As Long, lim As Long

    at = CLng(Val(APS_CfgLer("CEN_IDATUAL", "0")))
    lim = CLng(Val(APS_CfgLer("CEN_IDLIMITE", "0")))
    If at >= lim Then
        If Not Central_ReservarBloco() Then Exit Function
        at = CLng(Val(APS_CfgLer("CEN_IDATUAL", "0")))
        lim = CLng(Val(APS_CfgLer("CEN_IDLIMITE", "0")))
        If at >= lim Then Exit Function
    End If
    If consome Then
        at = at + 1
        APS_CfgGravar "CEN_IDATUAL", CStr(at)
        If CLng(Val(APS_CfgLer("CEN_IDATUAL", "0"))) <> at Then Exit Function
    Else
        at = at + 1
    End If
    Central_ProximoIdCentral = CStr(at)
End Function

' le PROXID central, soma CEN_LOTE_IDS com If-Match; repete em 412 (ate 5x). PREMISSA: 412 funciona (T-E1/T-P1).
Public Function Central_ReservarBloco() As Boolean
    Dim tent As Long, st As Long, corpo As String, etag As String, itemId As String
    Dim atual As Long, novo As Long, resp As String, et2 As String, a As Boolean

    On Error GoTo Falha
    For tent = 1 To 5
        corpo = Central_Http("GET", Central_UrlCfg() & "?expand=fields&$filter=fields/Chave eq 'PROXID'", "", "", st, etag)
        If st <> 200 Then mUltErro = "PROXID HTTP " & st: Exit Function
        Dim itens() As String, n As Long
        n = Central_JsonItens(Central_JsonGet(corpo, "value"), itens)
        If n <> 1 Then mUltErro = "PROXID: " & n & " itens": Exit Function
        itemId = Central_JsonGet(itens(1), "id")
        etag = Central_JsonGet(itens(1), "eTag")
        atual = CLng(Val(Central_JsonGet(Central_JsonGet(itens(1), "fields"), "Valor")))
        novo = atual + CEN_LOTE_IDS
        resp = Central_Http("PATCH", Central_UrlCfg() & "/" & itemId & "/fields", "{""Valor"":""" & CStr(novo) & """}", etag, st, et2)
        If st = 200 Then
            APS_CfgGravar "CEN_IDATUAL", CStr(atual)
            APS_CfgGravar "CEN_IDLIMITE", CStr(novo)
            Central_ReservarBloco = True
            Exit Function
        ElseIf st <> 412 Then
            mUltErro = "PROXID PATCH HTTP " & st
            Exit Function
        End If
    Next tent
    mUltErro = "PROXID: 5 conflitos seguidos"
    Exit Function
Falha:
    mUltErro = "reserva de IDs: " & Err.Description
End Function

'==========================================================
' TRANSPORTE (NAO TESTADO - depende de autenticacao autorizada pela TI)
'==========================================================
Private Function Central_UrlBase(ByVal lista As String) As String
    Central_UrlBase = "https://graph.microsoft.com/v1.0/sites/" & APS_CfgLer("CEN_SITEID", "") & "/lists/" & lista & "/items"
End Function

Public Function Central_UrlOps() As String
    Central_UrlOps = Central_UrlBase(APS_CfgLer("CEN_LISTA_OPS", ""))
End Function

Public Function Central_UrlCfg() As String
    Central_UrlCfg = Central_UrlBase(APS_CfgLer("CEN_LISTA_CFG", ""))
End Function

Public Function Central_UrlAud() As String
    Central_UrlAud = Central_UrlBase(APS_CfgLer("CEN_LISTA_AUD", ""))
End Function

' PONTO DE ENCAIXE DA AUTENTICACAO. Por padrao NAO HA metodo autorizado: devolve "" e registra o motivo.
' Quando a TI definir o metodo (app registrado/MSAL/helper aprovado), implemente SOMENTE aqui.
' Regras: token so em memoria; nunca gravar token/segredo em celula, nome definido ou log.
Private Function Central_Token() As String
    Central_Token = ""
    mUltErro = "autenticacao nao configurada (metodo autorizado pela TI pendente)"
End Function

' Chamada HTTP. Devolve o corpo; st = status HTTP (0 = falha de rede/autenticacao); etagResp = cabecalho ETag.
Public Function Central_Http(ByVal metodo As String, ByVal url As String, ByVal corpo As String, _
                             ByVal ifMatch As String, ByRef st As Long, ByRef etagResp As String) As String
    Dim h As Object, tk As String

    st = 0: etagResp = ""
    tk = Central_Token()
    If Len(tk) = 0 Then Exit Function

    On Error GoTo Falha
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.setTimeouts 5000, 5000, 15000, 30000
    h.Open metodo, url, False
    h.setRequestHeader "Authorization", "Bearer " & tk
    h.setRequestHeader "Content-Type", "application/json"
    h.setRequestHeader "Accept", "application/json"
    h.setRequestHeader "Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly"
    If Len(ifMatch) > 0 Then h.setRequestHeader "If-Match", ifMatch
    h.send corpo
    st = h.Status
    On Error Resume Next
    etagResp = h.getResponseHeader("ETag")
    On Error GoTo 0
    Central_Http = h.responseText
    tk = String$(Len(tk), "*")      ' nao deixa o token em variavel
    Exit Function
Falha:
    st = 0
    mUltErro = "rede: " & Err.Description
End Function

'==========================================================
' SINCRONIZACAO (chamada por APS_Tick; ver LEIA-ME para o gancho)
'==========================================================
Public Sub Central_Sincronizar(Optional ByVal forcar As Boolean = False)
    If Not Central_Conectado() Then Exit Sub
    If mEmSync Then Exit Sub
    If APS_Ocupado > 0 Then Exit Sub
    If Not forcar Then
        If mUltSync > 0 And (Now - mUltSync) * 86400# < Central_IntervaloSeg() - 5 Then Exit Sub
    End If

    mEmSync = True
    On Error GoTo Falha
    mUltErro = ""
    If Central_Enviar() Then Central_Puxar
    mUltSync = Now
    Application.StatusBar = Central_Estado()
    mEmSync = False
    Exit Sub
Falha:
    mUltErro = "sync: " & Err.Description
    mEmSync = False
    On Error Resume Next
    Application.StatusBar = Central_Estado()
End Sub

' Localiza item no mapa (ItemId/ETag) pelo ApsId. Devolve a linha do mapa ou 0.
Private Function Central_MapaLinha(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim u As Long, r As Long
    u = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    For r = 2 To u
        If CStr(ws.Cells(r, 8).Value) = id Then Central_MapaLinha = r: Exit Function
    Next r
End Function

Private Sub Central_MapaGravar(ByVal ws As Worksheet, ByVal id As String, ByVal itemId As String, ByVal etag As String, ByVal modUtc As String)
    Dim r As Long
    r = Central_MapaLinha(ws, id)
    If r = 0 Then r = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 8).Value = id
    ws.Cells(r, 9).Value = itemId
    ws.Cells(r, 10).Value = etag
    ws.Cells(r, 11).Value = modUtc
End Sub

' Envia pendentes. Devolve True se a fila foi esvaziada ou so restam pendencias que NAO impedem o pull.
' Sem transacao: itens aceitos permanecem aceitos mesmo se um posterior falhar.
Private Function Central_Enviar() As Boolean
    Dim ws As Worksheet, wo As Worksheet, ops() As tOperacao, n As Long, r As Long, i As Long
    Dim id As String, acao As String, chave As String, st As Long, resp As String, et As String
    Dim mr As Long, itemId As String, etag As String, corpo As String, est As Boolean, k As Long
    Dim iAchou As Long, tent As Long

    Central_Enviar = True
    Set ws = Central_AbaFila()
    If ws Is Nothing Then mUltErro = "aba APS_FILA ausente (rode Central_Instalar)": Central_Enviar = False: Exit Function
    If Central_FilaPendentes() = 0 Then Exit Function

    n = APS_LerOps(ops)
    est = APS_Liberar(ws)
    r = 2
    Do While r <= ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        id = CStr(ws.Cells(r, 1).Value)
        If Len(id) = 0 Then Exit Do
        acao = CStr(ws.Cells(r, 2).Value)
        chave = CStr(ws.Cells(r, 3).Value)
        mr = Central_MapaLinha(ws, id)

        iAchou = 0
        For i = 1 To n
            If Len(ops(i).tipoAux) = 0 Then
                If APS_Ig(ops(i).id, id) Then iAchou = i: Exit For
            End If
        Next i

        If mr = 0 And (acao = "EXCLUIR" Or iAchou = 0) Then
            ' nunca existiu no servidor: nao ha o que excluir
            Central_FilaRemover ws, r
        Else
        If acao = "EXCLUIR" Or iAchou = 0 Then
            ' exclusao logica: nunca apaga o item central
            corpo = "{""Excluido"":true}"
        Else
            corpo = Central_OpParaCampos(ops(iAchou), chave, False)
        End If

        If mr = 0 Then
            ' criar (ChaveCliente unico: reenvio apos timeout nao duplica; HTTP 409 = ja existe -> consulta pela chave)
            resp = Central_Http("POST", Central_UrlOps(), "{""fields"":" & corpo & "}", "", st, et)
            If st = 409 Then
                resp = Central_Http("GET", Central_UrlOps() & "?expand=fields&$filter=fields/ChaveCliente eq '" & chave & "'", "", "", st, et)
            End If
        Else
            itemId = CStr(ws.Cells(mr, 9).Value)
            etag = CStr(ws.Cells(mr, 10).Value)
            resp = Central_Http("PATCH", Central_UrlOps() & "/" & itemId & "/fields", corpo, etag, st, et)
        End If

        Select Case st
            Case 200, 201
                ' aceito: guarda ItemId/ETag devolvidos e tira da fila
                itemId = Central_JsonGet(resp, "id")
                If Len(itemId) = 0 And mr > 0 Then itemId = CStr(ws.Cells(mr, 9).Value)
                etag = Central_JsonGet(resp, "eTag")
                If Len(etag) = 0 Then etag = et
                Central_MapaGravar ws, id, itemId, etag, ""
                Central_FilaRemover ws, r
                Central_Audit IIf(acao = "EXCLUIR", "excluir", "atualizar"), id, "ok", "", chave
                ' nao avanca r: a linha r agora contem a ultima pendencia
            Case 412
                If Central_DecisaoConflitoEtag(acao) = "DESCARTAR" Then
                    Central_FilaRemover ws, r
                    Central_Audit "conflito", id, "status automatico descartado", "412", chave
                    k = k + 1
                Else
                    ws.Cells(r, 6).Value = "CONFLITO_ETAG: alterada por outra pessoa; revise antes de reenviar"
                    ws.Cells(r, 4).Value = CLng(Val(ws.Cells(r, 4).Value)) + 1
                    mUltErro = "conflito de edicao (" & id & ")"
                    Central_Audit "conflito", id, "412 mantido pendente", "", chave
                    r = r + 1
                End If
            Case 0, 401, 403, 404, 429, 503
                ws.Cells(r, 6).Value = "HTTP " & st
                ws.Cells(r, 4).Value = CLng(Val(ws.Cells(r, 4).Value)) + 1
                If Len(mUltErro) = 0 Then mUltErro = "envio HTTP " & st
                Central_Enviar = False           ' sem conexao/permissao: nao tenta o pull agora
                Exit Do
            Case Else
                ws.Cells(r, 6).Value = "HTTP " & st
                ws.Cells(r, 4).Value = CLng(Val(ws.Cells(r, 4).Value)) + 1
                r = r + 1
        End Select
        End If
        tent = tent + 1
        If tent > 500 Then Exit Do
    Loop
    APS_Reproteger ws, est
End Function

' Le alteracoes remotas (janela >= cursor, dedupe por ETag), valida o conjunto e SO ENTAO aplica.
' O cursor so avanca com o resultado do pull.
Private Sub Central_Puxar()
    Dim ws As Worksheet, st As Long, url As String, corpo As String, et As String, itens() As String, n As Long
    Dim i As Long, campos As String, id As String, itemId As String, etag As String, modU As String, mx As String
    Dim o As tOperacao, setupSeg As Double, excl As Boolean, motivo As String, mr As Long, est As Boolean
    Dim aplic As Long, rej As Long, pend As Boolean, wo As Worksheet, estO As Boolean

    Set ws = Central_AbaFila()
    If ws Is Nothing Then Exit Sub
    If Len(mCursor) = 0 Then mCursor = APS_CfgLer("CEN_CURSOR", "")

    url = Central_UrlOps() & "?expand=fields&$top=200"
    If Len(mCursor) > 0 Then url = url & "&$filter=fields/Modified ge '" & mCursor & "'"
    corpo = Central_Http("GET", url, "", "", st, et)
    If st <> 200 Then mUltErro = "pull HTTP " & st: Exit Sub

    n = Central_JsonItens(Central_JsonGet(corpo, "value"), itens)
    mx = mCursor
    Set wo = APS_Aba(APS_ABA_OPS)
    If wo Is Nothing Then Exit Sub

    est = APS_Liberar(ws)
    estO = APS_Liberar(wo)
    mAplicando = True
    APS_Ocupado = APS_Ocupado + 1
    On Error GoTo Falha

    For i = 1 To n
        campos = Central_JsonGet(itens(i), "fields")
        itemId = Central_JsonGet(itens(i), "id")
        etag = Central_JsonGet(itens(i), "eTag")
        modU = Central_JsonGet(campos, "Modified")
        If Len(modU) = 0 Then modU = Central_JsonGet(itens(i), "lastModifiedDateTime")
        If modU > mx Then mx = modU

        motivo = Central_CamposParaOp(campos, o, setupSeg, excl)
        If Len(motivo) > 0 Then
            rej = rej + 1
            Central_Audit "conflito", Central_JsonGet(campos, "ApsId"), "item remoto rejeitado", motivo, ""
        Else
            id = o.id
            mr = Central_MapaLinha(ws, id)
            If mr > 0 And CStr(ws.Cells(mr, 10).Value) = etag Then
                ' mesmo ETag ja conhecido (janela de sobreposicao): nada a fazer
            ElseIf Central_FilaLinha(ws, id) > 0 Then
                ' ha alteracao local pendente para este item: NAO sobrescreve
                mUltErro = "item " & id & " alterado em outro lugar e tambem localmente (pendente)"
                Central_Audit "conflito", id, "pull nao sobrescreveu pendente local", "", ""
            Else
                If Central_AplicarNaPlanilha(wo, o, setupSeg, excl) Then
                    Central_MapaGravar ws, id, itemId, etag, modU
                    aplic = aplic + 1
                Else
                    rej = rej + 1
                End If
            End If
        End If
    Next i

    ' cursor avanca SOMENTE aqui, com o resultado do pull (nunca com a resposta de uma gravacao propria)
    If mx <> mCursor Then
        mCursor = mx
        APS_CfgGravar "CEN_CURSOR", mx
    End If

Saida:
    On Error Resume Next
    APS_Ocupado = APS_Ocupado - 1
    mAplicando = False
    APS_Reproteger wo, estO
    APS_Reproteger ws, est
    If aplic > 0 Then APS_Atualizar True
    Exit Sub
Falha:
    mUltErro = "pull: " & Err.Description
    Resume Saida
End Sub

' Escreve um item remoto na planilha local (somente valores em colunas existentes).
Private Function Central_AplicarNaPlanilha(ByVal wo As Worksheet, ByRef o As tOperacao, ByVal setupSeg As Double, ByVal excl As Boolean) As Boolean
    Dim cId As Long, r As Long, u As Long

    cId = APS_Col(wo, H_ID())
    If cId = 0 Then mUltErro = "coluna ID ausente na planilha (estrutura nao e alterada em modo conectado)": Exit Function
    u = wo.Cells(wo.Rows.Count, cId).End(xlUp).Row
    For r = 2 To u
        If CStr(wo.Cells(r, cId).Value) = o.id Then o.linha = r: Exit For
    Next r

    If excl Then
        If o.linha >= 2 Then APS_LimparLinha wo, o.linha
        Central_AplicarNaPlanilha = True
        Exit Function
    End If

    APS_GravarOp o
    If o.linha >= 2 Then APS_GravarSetupDaOp wo, o.linha, Central_SegParaDias(setupSeg)
    Central_AplicarNaPlanilha = (o.linha >= 2)
End Function

' Chamada em Workbook_BeforeClose: avisa se ha alteracoes ainda nao enviadas (nao impede o fechamento)
Public Sub Central_AoFechar()
    Dim p As Long
    If Not Central_Conectado() Then Exit Sub
    On Error Resume Next
    p = Central_FilaPendentes()
    If p > 0 Then
        MsgBox "Aten" & ChrW(231) & ChrW(227) & "o: " & p & " altera" & ChrW(231) & ChrW(227) & "o(" & ChrW(245) & "es) ainda n" & ChrW(227) & "o foi(ram) enviada(s) " & ChrW(224) & " lista central." & vbCrLf & _
               "Elas continuam salvas neste arquivo e ser" & ChrW(227) & "o enviadas na pr" & ChrW(243) & "xima abertura com conex" & ChrW(227) & "o.", vbExclamation, APS_TITULO
    End If
End Sub

'==========================================================
' AUDITORIA (melhor esforco; nunca bloqueia; sem segredos)
'==========================================================
Public Sub Central_Audit(ByVal acao As String, ByVal id As String, ByVal resultado As String, ByVal detalhe As String, ByVal chave As String)
    Dim st As Long, et As String, corpo As String

    On Error Resume Next
    corpo = "{""fields"":{""Title"":""" & Central_JsonEsc(acao) & """,""Acao"":""" & Central_JsonEsc(acao) & """," & _
            """ApsId"":""" & Central_JsonEsc(id) & """,""Resultado"":""" & Central_JsonEsc(resultado) & """," & _
            """Detalhe"":""" & Central_JsonEsc(Left$(detalhe, 250)) & """,""ChaveCliente"":""" & Central_JsonEsc(chave) & """}}"
    Central_Http "POST", Central_UrlAud(), corpo, "", st, et
End Sub

'==========================================================
' AUTOTESTE (rodar no Excel: Alt+F8 > Central_AutoTeste). So partes puras. NAO testa SharePoint.
'==========================================================
Private Sub T_Ok(ByRef rel As String, ByRef falhas As Long, ByVal nome As String, ByVal cond As Boolean)
    If cond Then
        rel = rel & "OK    " & nome & vbCrLf
    Else
        rel = rel & "FALHA " & nome & vbCrLf
        falhas = falhas + 1
    End If
End Sub

Public Sub Central_AutoTeste()
    Dim rel As String, f As Long, s As Double, iso As String, o As tOperacao, o2 As tOperacao
    Dim js As String, ss As Double, ex As Boolean, it() As String, n As Long, a As Boolean

    ' tempo: 2026-10-22 08:00 local = 11:00Z
    s = CDbl(DateSerial(2026, 10, 22)) + TimeSerial(8, 0, 0)
    iso = Central_SerialParaIso(s)
    T_Ok rel, f, "tempo local->UTC (" & iso & ")", (iso = "2026-10-22T11:00:00Z")
    T_Ok rel, f, "tempo ida e volta", (Abs(Central_IsoParaSerial(iso) - s) < 0.00000002)
    s = CDbl(DateSerial(2026, 12, 31)) + TimeSerial(22, 30, 0)
    T_Ok rel, f, "virada de dia/ano", (Central_SerialParaIso(s) = "2027-01-01T01:30:00Z")
    T_Ok rel, f, "iso invalido -> 0", (Central_IsoParaSerial("lixo") = 0)
    T_Ok rel, f, "segundos exatos", (Central_DiasParaSeg(Central_SegParaDias(5400)) = 5400)

    ' JSON
    js = "{""a"":""x\ny\u00e9"",""b"":12.5,""c"":null,""d"":{""e"":[1,2]},""f"":[{""id"":""1""},{""id"":""2""}]}"
    T_Ok rel, f, "json string+escape", (Central_JsonGet(js, "a") = "x" & vbLf & "y" & ChrW(233))
    T_Ok rel, f, "json numero", (Central_JsonGet(js, "b") = "12.5")
    T_Ok rel, f, "json null", (Central_JsonGet(js, "c", a) = "" And a)
    T_Ok rel, f, "json ausente", (Central_JsonGet(js, "zz", a) = "" And Not a)
    T_Ok rel, f, "json objeto bruto", (Central_JsonGet(Central_JsonGet(js, "d"), "e") = "[1,2]")
    n = Central_JsonItens(Central_JsonGet(js, "f"), it)
    T_Ok rel, f, "json array (2 itens)", (n = 2 And Central_JsonGet(it(2), "id") = "2")
    T_Ok rel, f, "json escape acentos", (Central_JsonEsc("a" & ChrW(231) & """") = "a\u00e7\""")

    ' mapeamento ida e volta
    o.id = "7": o.codigo = "123": o.Produto = "Prod " & ChrW(227) & "o": o.lote = "L1": o.ordem = "OP9"
    o.maquina = "FETTE 2090": o.ini = CDbl(DateSerial(2026, 10, 22)) + TimeSerial(8, 0, 0)
    o.Fim = o.ini + TimeSerial(10, 0, 0): o.dur = TimeSerial(10, 0, 0): o.status = APS_ST_PLANEJADA
    o.caixas = 100: o.comp = 3000: o.oee = 0.8: o.velBase = 500: o.obs = "obs ""q"""
    js = Central_OpParaCampos(o, "APS-T", False)
    T_Ok rel, f, "mapeamento: sem erro", (Central_CamposParaOp(js, o2, ss, ex) = "")
    T_Ok rel, f, "mapeamento: campos", (o2.id = "7" And o2.maquina = "FETTE 2090" And o2.lote = "L1" And o2.status = APS_ST_PLANEJADA And o2.obs = o.obs And o2.Produto = o.Produto)
    T_Ok rel, f, "mapeamento: datas", (Abs(o2.ini - o.ini) < 0.00000002 And Abs(o2.Fim - o.Fim) < 0.00000002)
    T_Ok rel, f, "mapeamento: rejeita Fim<Inicio", (Len(Central_CamposParaOp(Replace(js, "2026-10-22T21:00:00Z", "2026-10-22T01:00:00Z"), o2, ss, ex)) > 0)
    T_Ok rel, f, "mapeamento: rejeita status", (Len(Central_CamposParaOp(Replace(js, """Planejada""", """Xyz"""), o2, ss, ex)) > 0)

    ' sobreposicao
    T_Ok rel, f, "sobreposicao parcial", Central_Sobrepoe(1, 3, 2, 4)
    T_Ok rel, f, "encostar nao sobrepoe", Not Central_Sobrepoe(1, 2, 2, 3)
    T_Ok rel, f, "conflito 412: status auto descarta", (Central_DecisaoConflitoEtag("STATUSAUTO") = "DESCARTAR")
    T_Ok rel, f, "conflito 412: edicao do usuario mantem", (Central_DecisaoConflitoEtag("UPSERT") = "MANTER")

    MsgBox "Autoteste Central (somente logica local; SharePoint NAO testado)" & vbCrLf & vbCrLf & rel & vbCrLf & _
           IIf(f = 0, "Todos passaram.", f & " falha(s)."), IIf(f = 0, vbInformation, vbExclamation), APS_TITULO
End Sub
