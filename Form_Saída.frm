VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Saída 
   Caption         =   "Saída"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Saída.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Saída"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adicionar_Click()
 Comprador = UCase(Comprador.Value)
Call txparalv
End Sub
Private Sub Saida_Click()
Saída.Show
Unload Me
End Sub
Private Sub Adicionar2_Click()
Cadastrar.Show
End Sub
Private Sub Quantidade_Change()
If Quantidade.Value <> "" Then
If Not IsNumeric(Quantidade) Then
MsgBox "Somente números", vbInformation, "Erro"
Quantidade.Text = ""
Exit Sub
End If
End If
End Sub
Private Sub Remover_Click()
Call apagar
End Sub
Private Sub Salvar_Click()
Call adicionarlvbd
Comprador.Value = ""
Call Codigo1
End Sub
Private Sub UserForm_Initialize()
Call Codigo1
Call Propriedadeslv2
Nome1.Enabled = False
End Sub
Sub txparalv()
If Nome1.Value <> "" And Data.Value <> "" And Quantidade.Value <> "" And valor.Value <> "" And Comprador.Value <> "" Then
Dim Data1 As Date
On Error GoTo erro
Data1 = Data
Dim i As Integer
Dim soma As Double
If Me.ListBox1.ListCount = 0 Then
    i = 0
Else
    i = Me.ListBox1.ListCount
End If
Me.ListBox1.AddItem Nome1.Value
ListBox1.List(i, 1) = Quantidade.Value
ListBox1.List(i, 2) = FormatNumber(valor.Text, 2)
ListBox1.List(i, 3) = FormatNumber(ListBox1.List(i, 1) * ListBox1.List(i, 2), 2)
ListBox1.List(i, 4) = Data1
ListBox1.List(i, 5) = Comprador.Value
soma = soma + ListBox1.List(i, 3)
Nome1.Value = ""
Quantidade.Value = ""
valor.Value = ""
Else
MsgBox "Preencha todos os campos", vbInformation, "Erro"
End If
Exit Sub
erro:
MsgBox "Digite uma data válida", vbInformation, "Erro"
Data.Value = ""
Exit Sub
End Sub
Private Sub Propriedadeslv2()
With ListBox1
.Clear
.ColumnWidths = "115;90;80;90;75;80"
.ColumnCount = 6
.ListStyle = fmListStylePlain
.AddItem
.List(0, 0) = "NOME"
.List(0, 1) = "QUANTIDADE"
.List(0, 2) = "PREÇO UND"
.List(0, 3) = "PREÇO TOTAL"
.List(0, 4) = "DATA"
.List(0, 5) = "COMPRADOR"
End With
End Sub
Private Sub adicionarlvbd()
If ListBox1.ListCount > 1 Then
Plan2.Select
  Dim Item As Double
   Dim linha As Integer
With Plan2
linha = Cells(Rows.Count, "b").End(3).Row + 1
 For Item = 1 To ListBox1.ListCount - 1
.Cells(linha, 2) = Codigo.Caption
.Cells(linha, 3) = ListBox1.List(Item, 0)
.Cells(linha, 4) = (ListBox1.List(Item, 1))
.Cells(linha, 5) = CDbl(ListBox1.List(Item, 2))
.Cells(linha, 6) = CDbl(ListBox1.List(Item, 3))
.Cells(linha, 7) = CDate(ListBox1.List(Item, 4))
.Cells(linha, 8) = ListBox1.List(Item, 5)
linha = linha + 1
Next
MsgBox "Informações Salvas Com Sucesso.", vbInformation, "Salvar"
ListBox1.Clear
End With
Else
MsgBox "Adicione um Produto", vbInformation, "Erro"
End If
End Sub
Private Sub Codigo1()
If Plan2.Range("b4").Value <> "" Then
Dim cd As Integer
Plan2.Select
Plan2.Range("b4").Select
Do
If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True
cd = ActiveCell.Offset(-1, 0).Value
Codigo.Caption = cd + 1
Else
Codigo.Caption = 1
End If
End Sub
Sub apagar()
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then
Dim Nome As String
Nome = ListBox1.List(i, 0)
ListBox1.RemoveItem (i)
End If
Next i
End Sub
Private Sub valor_Change()
If valor.Value <> "" Then
If Not IsNumeric(valor) Then
MsgBox "Somente números", vbInformation, "Erro"
valor.Text = ""
Exit Sub
End If
End If
End Sub
