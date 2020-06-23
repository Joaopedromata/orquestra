VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Editar 
   Caption         =   "Editar"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Editar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Consulta_Change()
Call Carregarlv
End Sub
Private Sub Excluir_Click()
Resp = MsgBox("Deseja Excluir Esses Dados?", vbYesNo, "Excluir")
If Resp = vbYes Then
If Consulta.Value = "Entrada" Then

For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then
Nome = ListBox1.List(i, 0)
ListBox1.RemoveItem (i)
Plan1.Select
Plan1.Range("c4").Select
With Worksheets("Plan1").Range("C:C")
Set C = .Find(Nome, LookIn:=xlValues, lookat:=xlWhole)
If Not C Is Nothing Then
C.Activate
Selection.EntireRow.Delete
End If
End With
End If
Next i
ElseIf Consulta.Value = "Saída" Then
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then

Nome = ListBox1.List(i, 0)
ListBox1.RemoveItem (i)
Plan2.Select
Plan2.Range("c4").Select
With Worksheets("Plan2").Range("C:C")
Set C = .Find(Nome, LookIn:=xlValues, lookat:=xlWhole)
If Not C Is Nothing Then
C.Activate
Selection.EntireRow.Delete
End If
End With
End If
Next i

ElseIf Consulta.Value = "Produtos Cadastrados" Then
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then

Nome = ListBox1.List(i, 0)
ListBox1.RemoveItem (i)
Plan3.Select
Plan3.Range("c4").Select
With Worksheets("Plan3").Range("C:C")
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
Private Sub UserForm_Initialize()
With Consulta
.AddItem ("Entrada")
.AddItem ("Saída")
.AddItem ("Produtos Cadastrados")
End With
End Sub
Sub Carregarlv()
Dim ultimalinha As Long
Dim linha As Integer
If Consulta.Value = "Entrada" Then
ListBox1.Clear
Call propriedadeslb1
ultimalinha = Plan1.Range("c1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan1.Range("c" & linha)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan1.Cells(linha, "d")
ListBox1.List(ListBox1.ListCount - 1, 2) = FormatNumber(Plan1.Cells(linha, "e"), 2)
ListBox1.List(ListBox1.ListCount - 1, 3) = FormatNumber(Plan1.Cells(linha, "f"), 2)
ListBox1.List(ListBox1.ListCount - 1, 4) = Plan1.Cells(linha, "g")
ListBox1.List(ListBox1.ListCount - 1, 5) = Plan1.Cells(linha, "h")
Next









ElseIf Consulta.Value = "Saída" Then
ListBox1.Clear
Call propriedadeslb2
ultimalinha = Plan2.Range("c1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan2.Range("c" & linha)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan2.Cells(linha, "d")
ListBox1.List(ListBox1.ListCount - 1, 2) = FormatNumber(Plan2.Cells(linha, "e"), 2)
ListBox1.List(ListBox1.ListCount - 1, 3) = FormatNumber(Plan2.Cells(linha, "f"), 2)
ListBox1.List(ListBox1.ListCount - 1, 4) = Plan2.Cells(linha, "g")
ListBox1.List(ListBox1.ListCount - 1, 5) = Plan2.Cells(linha, "h")
Next
ElseIf Consulta.Value = "Produtos Cadastrados" Then
ListBox1.Clear
Call propriedadeslb3
ultimalinha = Plan3.Range("c1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan3.Range("c" & linha)
Next


ElseIf Consulta.Value = "Produtos Cadastrados" Then
ListBox1.Clear
Call propriedadeslb3
ultimalinha = Plan3.Range("C1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan3.Range("C" & linha)
Next


















End If

End Sub
Sub propriedadeslb1()
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
.List(0, 5) = "FORNECEDOR"
End With
End Sub
Sub propriedadeslb2()
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
Sub propriedadeslb3()
With ListBox1
.Clear
.ColumnWidths = "330"
.ColumnCount = 1
.ListStyle = fmListStylePlain
.AddItem
.List(0, 0) = "NOME"
End With
End Sub























