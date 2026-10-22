Attribute VB_Name = "modAPS_Diagnostico"
Option Explicit

'==========================================================
' APS PURAN - Diagnostico de consistencia (SOMENTE LEITURA)
' Nao grava nada, nao move nada e nao altera nenhuma regra. Serve para conferir se o que esta em 02_Operacoes
' e coerente e se os dados necessarios ao Painel existem.
'
' Dois pontos de entrada:
'   APS_Diagnostico         (Alt+F8) - mostra os alertas numa caixa de mensagem (comportamento original, + novas checagens)
'   APS_DiagnosticoColetar  - devolve os mesmos alertas em uma lista (usada pelo modAPS_Painel; nenhum codigo duplicado)
' Cada alerta e RASTREAVEL: tipo, gravidade, linha da planilha, OP/produto/lote/maquina e o texto do achado.
' Nenhum dado e corrigido aqui: o diagnostico apenas aponta.
'==========================================================

' gravidade: 1 = alta, 2 = media, 3 = baixa
Public Type tAlerta
    cod As String        ' FIM, ID, SOBREP, IGNORADA, SEMDUR, SEMQTD, STATUS, ORFAO, MAQ, SEMLOTE
    sev As Long
    linha As Long        ' linha em 02_Operacoes (0 = nao se aplica)
    ref As String        ' OP / produto / lote / maquina
    ini As Double
    fim As Double
    texto As String
End Type

Private Function DiagNome(ByRef o As tOperacao) As String
    If Len(o.tipoAux) > 0 Then
        DiagNome = o.rotAux & " de " & o.Produto & " lote " & o.lote & " (" & o.maquina & ")"
    Else
        DiagNome = "OP " & o.ordem & " " & o.Produto & " lote " & o.lote & " (" & o.maquina & ")"
    End If
End Function

Private Sub DiagPush(ByRef al() As tAlerta, ByRef c As Long, ByVal cod As String, ByVal sev As Long, _
                     ByRef o As tOperacao, ByVal texto As String)
    c = c + 1
    If c > UBound(al) Then ReDim Preserve al(1 To UBound(al) + 200)
    al(c).cod = cod
    al(c).sev = sev
    al(c).linha = o.linha
    al(c).ref = DiagNome(o)
    al(c).ini = o.ini
    al(c).fim = o.fim
    al(c).texto = texto
End Sub

Private Function DiagMaquinaCadastrada(ByVal nome As String) As Boolean
    Dim m() As String, k As Long, i As Long
    k = APS_Maquinas(m)
    For i = 1 To k
        If APS_Ig(m(i), nome) Then DiagMaquinaCadastrada = True: Exit Function
    Next i
End Function

Private Function DiagStatusValido(ByVal st As String) As Boolean
    Dim v As Variant, i As Long
    v = APS_ListaStatus()
    For i = LBound(v) To UBound(v)
        If APS_Ig(CStr(v(i)), st) Then DiagStatusValido = True: Exit Function
    Next i
End Function

'----------------------------------------------------------
' Coletor: devolve em al(1..retorno) todos os alertas. Retorno -1 = nao foi possivel ler 02_Operacoes.
'----------------------------------------------------------
Public Function APS_DiagnosticoColetar(ByRef al() As tAlerta) As Long
    Dim ops() As tOperacao, n As Long, i As Long, j As Long, c As Long
    Dim fimEsp As Double, parceiros As Boolean, achou As Boolean, ign As Long

    ReDim al(1 To 200)
    On Error GoTo Falha
    APS_CalRecarregar
    n = APS_LerOps(ops)
    ign = APS_OpsIgnoradas
    If n < 0 Then
        APS_DiagnosticoColetar = -1
        Exit Function
    End If

    ' 0) Linhas que o sistema ignora (sem data de inicio valida ou sem maquina): nao entram em nada
    If ign > 0 Then
        c = c + 1
        al(c).cod = "IGNORADA"
        al(c).sev = 2
        al(c).texto = ign & " linha(s) de 02_Operacoes sem data de in" & ChrW(237) & "cio v" & ChrW(225) & "lida ou sem m" & ChrW(225) & "quina: " & _
                      "n" & ChrW(227) & "o aparecem nos cards nem no Painel."
    End If

    For i = 1 To n
        If Len(ops(i).tipoAux) = 0 And APS_Ativa(ops(i)) Then
            ' 1) Fim coerente com Inicio + tempo programado (duracao + atraso) pelo calendario
            If Not ops(i).SemDur Then
                fimEsp = APS_AdicionarTrabalho(ops(i).ini, ops(i).dur)
                If fimEsp >= 0 Then
                    If Abs(fimEsp - ops(i).fim) > 1# / 1440# + APS_EPS Then
                        DiagPush al, c, "FIM", 1, ops(i), "Fim gravado " & Format$(ops(i).fim, "dd/mm hh:nn") & " <> esperado " & _
                            Format$(fimEsp, "dd/mm hh:nn") & " (In" & ChrW(237) & "cio " & Format$(ops(i).ini, "dd/mm hh:nn") & " + dura" & ChrW(231) & ChrW(227) & "o programada)."
                    End If
                End If
            Else
                ' 5) sem duracao e sem fim
                DiagPush al, c, "SEMDUR", 2, ops(i), "Sem dura" & ChrW(231) & ChrW(227) & "o e sem data de fim: o sistema assume 1 h apenas para exibir o card."
            End If
            ' 6) quantidade
            If APS_CompDaOp(ops(i)) <= APS_EPS Then
                DiagPush al, c, "SEMQTD", 3, ops(i), "Sem quantidade do lote (comprimidos): a ordem entra no Painel com 0 comprimidos."
            End If
            ' 7) status fora da lista
            If Not DiagStatusValido(ops(i).status) Then
                DiagPush al, c, "STATUS", 2, ops(i), "Status '" & ops(i).status & "' n" & ChrW(227) & "o " & ChrW(233) & " um status v" & ChrW(225) & "lido."
            End If
            ' 8) maquina fora do cadastro
            If Not DiagMaquinaCadastrada(ops(i).maquina) Then
                DiagPush al, c, "MAQ", 3, ops(i), "M" & ChrW(225) & "quina '" & ops(i).maquina & "' n" & ChrW(227) & "o consta no cadastro de m" & ChrW(225) & "quinas."
            End If
            ' 9) lote ausente
            If Len(Trim$(ops(i).lote)) = 0 Then
                DiagPush al, c, "SEMLOTE", 3, ops(i), "Lote n" & ChrW(227) & "o informado."
            End If
        End If
    Next i

    ' 2) IDs repetidos
    For i = 1 To n - 1
        For j = i + 1 To n
            If Len(ops(i).id) > 0 And APS_Ig(ops(i).id, ops(j).id) Then
                DiagPush al, c, "ID", 1, ops(i), "ID repetido (" & ops(i).id & "): " & DiagNome(ops(i)) & " e " & DiagNome(ops(j)) & "."
            End If
        Next j
    Next i

    ' 3) Sobreposicao na mesma maquina (producao e o proprio Setup/Limpeza nao contam entre si)
    For i = 1 To n - 1
        For j = i + 1 To n
            If APS_Ativa(ops(i)) And APS_Ativa(ops(j)) Then
                If APS_Ig(ops(i).maquina, ops(j).maquina) Then
                    parceiros = (Len(ops(i).idPai) > 0 And APS_Ig(ops(i).idPai, ops(j).id)) Or _
                                (Len(ops(j).idPai) > 0 And APS_Ig(ops(j).idPai, ops(i).id))
                    If Not parceiros Then
                        If APS_Sobrepoe(ops(i), ops(j)) Then
                            DiagPush al, c, "SOBREP", 1, ops(i), "Sobreposi" & ChrW(231) & ChrW(227) & "o em " & ops(i).maquina & ": " & DiagNome(ops(i)) & " (" & Format$(ops(i).ini, "dd/mm hh:nn") & " a " & _
                                Format$(ops(i).fim, "dd/mm hh:nn") & ") e " & DiagNome(ops(j)) & " (" & Format$(ops(j).ini, "dd/mm hh:nn") & " a " & Format$(ops(j).fim, "dd/mm hh:nn") & ")."
                        End If
                    End If
                End If
            End If
        Next j
    Next i

    ' 4) Setup/Limpeza sem producao vinculada
    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 Then
            achou = False
            If Len(Trim$(ops(i).idPai)) > 0 Then
                For j = 1 To n
                    If Len(ops(j).tipoAux) = 0 Then
                        If APS_Ig(ops(j).id, ops(i).idPai) Then achou = True: Exit For
                    End If
                Next j
            End If
            If Not achou Then
                DiagPush al, c, "ORFAO", 1, ops(i), DiagNome(ops(i)) & " n" & ChrW(227) & "o est" & ChrW(225) & " vinculado a nenhuma produ" & ChrW(231) & ChrW(227) & "o."
            End If
        End If
    Next i

    APS_DiagnosticoColetar = c
    Exit Function

Falha:
    APS_DiagnosticoColetar = -1
End Function

'----------------------------------------------------------
' Alt+F8 > APS_Diagnostico
'----------------------------------------------------------
Public Sub APS_Diagnostico()
    Dim al() As tAlerta, c As Long, i As Long, t As String, k As Long, sv As Long
    Dim n1 As Long, n2 As Long, n3 As Long

    c = APS_DiagnosticoColetar(al)
    If c < 0 Then
        MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel executar o diagn" & ChrW(243) & "stico (verifique a aba 02_Operacoes).", vbExclamation, APS_TITULO
        Exit Sub
    End If

    If c = 0 Then
        MsgBox "Nenhuma inconsist" & ChrW(234) & "ncia encontrada (Setup/Limpeza inclu" & ChrW(237) & "dos).", vbInformation, APS_TITULO
        Exit Sub
    End If

    For i = 1 To c
        Select Case al(i).sev
            Case 1: n1 = n1 + 1
            Case 2: n2 = n2 + 1
            Case Else: n3 = n3 + 1
        End Select
    Next i

    ' lista as mais graves primeiro (gravidade 1, depois 2, depois 3), no maximo 15
    For sv = 1 To 3
        For i = 1 To c
            If al(i).sev = sv And k < 15 Then
                k = k + 1
                t = t & "- [" & Choose(al(i).sev, "alta", "m" & ChrW(233) & "dia", "baixa") & "] " & al(i).texto
                If al(i).linha > 0 Then t = t & " (linha " & al(i).linha & ")"
                t = t & vbCrLf
            End If
        Next i
    Next sv
    If c > k Then t = t & "... e mais " & (c - k) & "." & vbCrLf

    MsgBox c & " alerta(s): " & n1 & " alta(s), " & n2 & " m" & ChrW(233) & "dia(s), " & n3 & " baixa(s). Nada foi alterado:" & vbCrLf & vbCrLf & t, vbExclamation, APS_TITULO
End Sub
