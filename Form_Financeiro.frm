VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Financeiro 
   Caption         =   "Lançamentos Financeiros"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Financeiro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Financeiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Excluir_Click()
Call apagar
End Sub
Sub apagar()
Dim Nome As String
Resp = MsgBox("Deseja Excluir Esses Dados?", vbYesNo, "Excluir")
If Resp = vbYes Then

If Selecione.Value = "Ganhos" Then
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then

Nome = ListBox1.List(i, 1)
ListBox1.RemoveItem (i)
Plan4.Select
Plan4.Range("c4").Select
With Worksheets("Plan4").Range("C:C")
Set C = .Find(Nome, LookIn:=xlValues, lookat:=xlWhole)
If Not C Is Nothing Then
C.Activate
Selection.EntireRow.Delete
End If
End With
End If
Next i


ElseIf Selecione.Value = "Gastos" Then
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then

Nome = ListBox1.List(i, 1)
ListBox1.RemoveItem (i)
Plan5.Select
Plan5.Range("c4").Select
With Worksheets("Plan5").Range("C:C")
Set C = .Find(Nome, LookIn:=xlValues, lookat:=xlWhole)
If Not C Is Nothing Then
C.Activate
Selection.EntireRow.Delete
End If
End With
End If
Next i
End If


End If



End Sub

Private Sub Salvar_Click()
Descricao = UCase(Descricao.Value)
Call Lancamentos_Financeiros
If Selecione.Value <> Empty And Descricao.Text <> "" And Data.Text <> "" And valor.Text <> "" Then
Call Carregarlv
End If
Descricao.Value = ""
valor.Value = ""
End Sub
Private Sub Selecione_Change()
Call Carregarlv
End Sub
Private Sub UserForm_Initialize()
Call propriedadeslb1
With Selecione
.AddItem ("Gastos")
.AddItem ("Ganhos")
End With
End Sub
Sub Lancamentos_Financeiros()
If Selecione.Value <> Empty And Descricao.Text <> "" And Data.Text <> "" And valor.Text <> "" Then
On Error GoTo erro
Dim Data1 As Date
Data1 = Data.Value
If Selecione.Value = "Gastos" Then
Sheets("plan5").Select
Range("b4").Select
 Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True
 ActiveCell.Offset(0, 0).Value = CDbl(valor.Text)
ActiveCell.Offset(0, 1).Value = Descricao.Text
ActiveCell.Offset(0, 2).Value = Data1
ElseIf Selecione.Value = "Ganhos" Then
Sheets("plan4").Select
Range("b4").Select
 Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True
 ActiveCell.Offset(0, 0).Value = CDbl(valor.Text)
ActiveCell.Offset(0, 1).Value = Descricao.Text
ActiveCell.Offset(0, 2).Value = Data1
End If
MsgBox "Salvo com Sucesso", vbInformation, "Cadastro"
Else
MsgBox "Preencha todos os campos", vbInformation, "Erro"
End If
Exit Sub
erro:
MsgBox "Digite uma data válida", vbInformation, "Erro"
Data.Value = ""
Exit Sub
End Sub
Sub Carregarlv()

Dim ultimalinha As Long
Dim linha As Integer
If Selecione.Value = "Ganhos" Then
ListBox1.Clear
Call propriedadeslb1
ultimalinha = Plan4.Range("B1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem FormatNumber(Plan4.Range("B" & linha), 2)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan4.Cells(linha, "C")
ListBox1.List(ListBox1.ListCount - 1, 2) = Plan4.Cells(linha, "D")
Next
ElseIf Selecione.Value = "Gastos" Then
ListBox1.Clear
Call propriedadeslb1
ultimalinha = Plan5.Range("B1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem FormatNumber(Plan5.Range("B" & linha), 2)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan5.Cells(linha, "C")
ListBox1.List(ListBox1.ListCount - 1, 2) = Plan5.Cells(linha, "D")
Next
End If
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
Sub propriedadeslb1()
With ListBox1
.Clear
.ColumnWidths = "80;200;80"
.ColumnCount = 3
.ListStyle = fmListStylePlain
.AddItem
.List(0, 0) = "VALOR"
.List(0, 1) = "NOME"
.List(0, 2) = "DATA"
End With
End Sub

















