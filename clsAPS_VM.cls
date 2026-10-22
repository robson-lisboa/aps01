VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAPS_VM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Eventos de aplicacao usados SOMENTE pela aba "Visao Geral Mensal":
' trocar Mes/Ano/Maquina/Status (B4:E4) recarrega a tabela; ativar a aba recarrega mantendo os filtros.
Public WithEvents App As Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    If Sh.Name <> APS_VM_NomeAba() Then Exit Sub
    If Intersect(Target, Sh.Range("B4:E4")) Is Nothing Then Exit Sub
    APS_VM_Atualizar
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not Sh.Parent Is ThisWorkbook Then Exit Sub
    If Sh.Name <> APS_VM_NomeAba() Then Exit Sub
    APS_VM_AoAtivar
End Sub
