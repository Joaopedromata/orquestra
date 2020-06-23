VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configuracoes 
   Caption         =   "Configurações"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Configuracoes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Configuracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ZerarEntrada()
Plan1.Select
Range("b4").Activate
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Activate
      Selection.EntireRow.Delete
       ActiveCell.Offset(-1, 0).Activate
         Selection.EntireRow.Delete
  End If
Loop Until IsEmpty(ActiveCell) = True
End Sub
Sub ZerarSaida()
Plan2.Select
Range("b4").Activate
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Activate
      Selection.EntireRow.Delete
       ActiveCell.Offset(-1, 0).Activate
         Selection.EntireRow.Delete
  End If
Loop Until IsEmpty(ActiveCell) = True
End Sub
Sub ZerarNome()
Plan3.Select
Range("c4").Activate
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Activate
      Selection.EntireRow.Delete
       ActiveCell.Offset(-1, 0).Activate
         Selection.EntireRow.Delete
  End If
Loop Until IsEmpty(ActiveCell) = True
End Sub
Sub ZerarGanhos()
Plan4.Select
Range("b4").Activate
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Activate
      Selection.EntireRow.Delete
       ActiveCell.Offset(-1, 0).Activate
         Selection.EntireRow.Delete
  End If
Loop Until IsEmpty(ActiveCell) = True
End Sub
Sub ZerarGastos()
Plan5.Select
Range("b4").Activate
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Activate
      Selection.EntireRow.Delete
       ActiveCell.Offset(-1, 0).Activate
         Selection.EntireRow.Delete
  End If
Loop Until IsEmpty(ActiveCell) = True
End Sub
Private Sub AbrirBD_Click()
Application.Visible = True
End Sub
Private Sub FecharBD_Click()
Application.Visible = False
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub Zerar_Click()
Resp = MsgBox("Deseja Formatar o Banco de Dados?", vbYesNo, "Excluir")
If Resp = vbYes Then
Call ZerarEntrada
Call ZerarSaida
Call ZerarNome
Call ZerarGastos
Call ZerarGanhos
End If
End Sub
