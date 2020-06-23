VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Consultar 
   Caption         =   "Consultar"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Consultar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Consultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adicionar2_Click()
Call filtrar2
Call soma
Call soma1

End Sub
Private Sub Consulta_Change()
Call Carregarlv
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Pesquisar_Change()
If Data1.Value <> "" And Data2.Value <> "" Then
Call filtrar2
Else
Dim linha As Integer
Dim linhalistbox As Integer
Dim valor_celula As String
If Consulta.Value = "Entrada" Then

linhalistbox = 0
linha = 4
ListBox1.Clear
propriedadeslb1
Plan1.Select
With Plan1
While .Cells(linha, 3).Value <> ""
valor_celula = Plan1.Cells(linha, 3).Value
If UCase(Left(valor_celula, Len(Pesquisar.Text))) = UCase(Pesquisar.Text) Then
With ListBox1
.AddItem
.List(linhalistbox + 1, 0) = Plan1.Cells(linha, 3)
.List(linhalistbox + 1, 1) = Plan1.Cells(linha, 4)
.List(linhalistbox + 1, 2) = FormatNumber(Plan1.Cells(linha, 5), 2)
.List(linhalistbox + 1, 3) = FormatNumber(Plan1.Cells(linha, 6), 2)
.List(linhalistbox + 1, 4) = Plan1.Cells(linha, 7)
.List(linhalistbox + 1, 5) = Plan1.Cells(linha, 8)
End With
linhalistbox = linhalistbox + 1
End If
linha = linha + 1
Wend
End With
End If
If Consulta.Value = "Saída" Then

linhalistbox = 0
linha = 4
ListBox1.Clear
propriedadeslb2
Plan2.Select
With Plan2
While .Cells(linha, 3).Value <> ""
valor_celula = Plan2.Cells(linha, 3).Value
If UCase(Left(valor_celula, Len(Pesquisar.Text))) = UCase(Pesquisar.Text) Then
With ListBox1
.AddItem
.List(linhalistbox + 1, 0) = Plan2.Cells(linha, 3)
.List(linhalistbox + 1, 1) = Plan2.Cells(linha, 4)
.List(linhalistbox + 1, 2) = FormatNumber(Plan2.Cells(linha, 5), 2)
.List(linhalistbox + 1, 3) = FormatNumber(Plan2.Cells(linha, 6), 2)
.List(linhalistbox + 1, 4) = Plan2.Cells(linha, 7)
.List(linhalistbox + 1, 5) = Plan2.Cells(linha, 8)
End With
linhalistbox = linhalistbox + 1
End If
linha = linha + 1
Wend
End With
End If
End If

Call soma
Call soma1
End Sub
Private Sub UserForm_Initialize()
With Consulta
.AddItem ("Entrada")
.AddItem ("Saída")
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
End If
Call soma
Call soma1
End Sub
Sub soma()
Dim soma As Long
For i = 1 To ListBox1.ListCount - 1
soma = soma + ListBox1.List(i, 1)
Next i
total.Caption = soma
End Sub
Sub filtrar2()
Dim linha As Integer
Dim linhalistview As Integer
Dim Data As Date
Dim Inicio As Date
Dim fim As Date
Dim valor_celula As String
If Consulta.Value = "Entrada" Then
If Data1.Value = "" Or Data2.Value = "" Then
MsgBox "Escolher data de Início e Fim.", vbInformation, "Erro"
Exit Sub
End If
On Error GoTo erro
Inicio = Data1.Value
fim = Data2.Value
linhalistview = 0
linha = 4
ListBox1.Clear
Call propriedadeslb1
Sheets("plan1").Select
With Sheets("plan1")
While .Cells(linha, 3).Value <> ""
valor_celula = .Cells(linha, 3).Value
Data = .Cells(linha, 7).Value
If Data >= Inicio And Data <= fim Then
If UCase(Left(valor_celula, Len(Pesquisar.Text))) = UCase(Pesquisar.Text) Then
ListBox1.AddItem Plan1.Range("c" & linha)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan1.Cells(linha, "d")
ListBox1.List(ListBox1.ListCount - 1, 2) = FormatNumber(Plan1.Cells(linha, "e"), 2)
ListBox1.List(ListBox1.ListCount - 1, 3) = FormatNumber(Plan1.Cells(linha, "f"), 2)
ListBox1.List(ListBox1.ListCount - 1, 4) = Plan1.Cells(linha, "g")
ListBox1.List(ListBox1.ListCount - 1, 5) = Plan1.Cells(linha, "h")
End If
End If
linha = linha + 1
Wend
End With
ElseIf Consulta.Value = "Saída" Then
On Error GoTo erro
Inicio = Data1.Value
fim = Data2.Value
linhalistview = 0
linha = 4
If Data1.Value = "" Or Data2.Value = "" Then
MsgBox "Escolher data de Início e Fim.", vbInformation, "Erro"
Exit Sub
End If
ListBox1.Clear
Call propriedadeslb2
Sheets("plan2").Select
With Sheets("plan2")
While .Cells(linha, 3).Value <> ""
valor_celula = .Cells(linha, 3).Value
Data = .Cells(linha, 7).Value
If Data >= Inicio And Data <= fim Then
If UCase(Left(valor_celula, Len(Pesquisar.Text))) = UCase(Pesquisar.Text) Then
ListBox1.AddItem Plan2.Range("c" & linha)
ListBox1.List(ListBox1.ListCount - 1, 1) = Plan2.Cells(linha, "d")
ListBox1.List(ListBox1.ListCount - 1, 2) = FormatNumber(Plan2.Cells(linha, "e"), 2)
ListBox1.List(ListBox1.ListCount - 1, 3) = FormatNumber(Plan2.Cells(linha, "f"), 2)
ListBox1.List(ListBox1.ListCount - 1, 4) = Plan2.Cells(linha, "g")
ListBox1.List(ListBox1.ListCount - 1, 5) = Plan2.Cells(linha, "h")
linhalistview = linhalistview + 1
End If
End If
linha = linha + 1
Wend
End With
End If
Exit Sub
erro:
MsgBox "Digite uma data válida", vbInformation, "Erro"
Data1.Value = ""
Data2.Value = ""
Exit Sub
Call soma
Call soma1

End Sub
Sub soma1()
Dim soma As Double
For i = 1 To ListBox1.ListCount - 1
soma = CDbl(soma) + CDbl(ListBox1.List(i, 3))
Next i
total1.Caption = "R$ " & FormatNumber(soma, 2)
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





