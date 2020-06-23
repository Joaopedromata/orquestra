VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menu"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cadastrar1_Click()
Cadastrar.Show
End Sub
Private Sub Configuracoes1_Click()
Configuracoes.Show
End Sub
Private Sub Consultar1_Click()
Consultar.Show
End Sub
Private Sub Editar1_Click()
Editar.Show
End Sub
Private Sub Entrada1_Click()
Entrada.Show
End Sub
Private Sub Estoque1_Click()
Estoque.Show

End Sub
Private Sub Financeiro1_Click()
financeiro.Show
End Sub
Private Sub Relatorio1_Click()
Relatório.Show
End Sub
Private Sub Saida1_Click()
Saída.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Workbooks(Application.ThisWorkbook.Name).Close savechanges:=True
ThisWorkbook.Save










End Sub
