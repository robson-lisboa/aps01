Attribute VB_Name = "modAPS_CentralMig"
Option Explicit
'==========================================================
' APS PURAN - MIGRACAO CONTROLADA PARA A LISTA CENTRAL (FASE 6)
' NAO executada em nenhum ambiente. So enfileira (nao envia) as producoes da planilha ATUAL.
' Regras: nunca migra automaticamente; uma planilha por vez (so o arquivo principal da Fase H);
' exige confirmacao do usuario; recusa se houver Setup legado, ID duplicado ou dado invalido
' (relatorio das divergencias, sem alterar nada). Usar SOMENTE com a lista de TESTE configurada.
'==========================================================
Public Sub Central_MigrarParaTeste()
    Dim ops() As tOperacao, n As Long, i As Long, j As Long, rel As String, ok As Long
    Dim leg As Long, dup As Long, inv As Long, r As VbMsgBoxResult

    If Not Central_Conectado() Then
        MsgBox "Ative o modo CONECTADO (Central_DefinirModo ""CONECTADO"") apontando para a lista de TESTE antes.", vbExclamation, APS_TITULO
        Exit Sub
    End If
    If Central_AbaFila() Is Nothing Then
        MsgBox "Rode Central_Instalar antes (cria a aba oculta APS_FILA).", vbExclamation, APS_TITULO
        Exit Sub
    End If

    n = APS_LerOps(ops)
    If n <= 0 Then MsgBox "Nenhuma operacao para migrar.", vbInformation, APS_TITULO: Exit Sub

    For i = 1 To n
        If Len(ops(i).tipoAux) > 0 Then
            leg = leg + 1: rel = rel & "Setup/Limpeza legado: ID " & ops(i).id & vbCrLf
        Else
            If Len(ops(i).id) = 0 Or Not IsNumeric(ops(i).id) Then
                inv = inv + 1: rel = rel & "ID ausente/invalido na linha " & ops(i).linha & vbCrLf
            End If
            For j = i + 1 To n
                If Len(ops(j).tipoAux) = 0 And Len(ops(i).id) > 0 Then
                    If APS_Ig(ops(i).id, ops(j).id) Then dup = dup + 1: rel = rel & "ID duplicado: " & ops(i).id & vbCrLf
                End If
            Next j
            If Len(Trim$(ops(i).maquina)) = 0 Or ops(i).ini <= 0 Then inv = inv + 1: rel = rel & "Dados invalidos (maquina/inicio), ID " & ops(i).id & vbCrLf
        End If
    Next i

    If leg + dup + inv > 0 Then
        MsgBox "Migracao RECUSADA (nada foi alterado). Divergencias:" & vbCrLf & rel & vbCrLf & _
               "Corrija na planilha (migracao de Setup legado e rotina administrativa separada) e rode de novo.", vbExclamation, APS_TITULO
        Exit Sub
    End If

    r = MsgBox("Enfileirar " & n & " operacao(oes) para envio a lista de TESTE?" & vbCrLf & "Nada e enviado agora: o envio ocorre na proxima sincronizacao.", vbYesNo + vbQuestion, APS_TITULO)
    If r <> vbYes Then Exit Sub

    For i = 1 To n
        If Central_FilaMarcar(ops(i).id, "UPSERT") Then ok = ok + 1
    Next i
    MsgBox ok & " de " & n & " enfileirada(s). Pendentes: " & Central_FilaPendentes() & ".", vbInformation, APS_TITULO
End Sub
