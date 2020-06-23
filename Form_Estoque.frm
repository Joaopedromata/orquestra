VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Estoque 
   Caption         =   "Estoque"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Estoque.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Frame1_Click()

End Sub

Private Sub Pesquisar_Change()
Call PesquisarLV
End Sub
Private Sub UserForm_Initialize()
propriedadeslb1
Call Carregarlv
Call ordemalfabetica
End Sub
Sub Carregarlv()
Dim ultimalinha As Long
Dim linha As Integer
Dim produtro As String
ultimalinha = Plan3.Range("C1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan3.Range("C" & linha)
produto = ListBox1.List(ListBox1.ListCount - 1, 0)
ListBox1.List(ListBox1.ListCount - 1, 1) = (WorksheetFunction.SumIfs(Plan1.Range("D4:D1000000"), Plan1.Range("C4:C1000000"), produto)) - (WorksheetFunction.SumIfs(Plan2.Range("d4:d1000000"), Plan2.Range("c4:c1000000"), produto))
Next
End Sub
Sub PesquisarLV()
Dim linha As Integer
Dim linhalistbox As Integer
Dim valor_celula As String
linhalistbox = 0
linha = 4
ListBox1.Clear
propriedadeslb1
Plan3.Select
With Plan3
While .Cells(linha, 3).Value <> ""
valor_celula = Plan3.Cells(linha, 3).Value
If UCase(Left(valor_celula, Len(Pesquisar.Text))) = UCase(Pesquisar.Text) Then
With ListBox1
.AddItem
.List(linhalistbox + 1, 0) = Plan3.Cells(linha, 3)
produto = ListBox1.List(ListBox1.ListCount - 1, 0)
ListBox1.List(ListBox1.ListCount - 1, 1) = (WorksheetFunction.SumIfs(Plan1.Range("D4:D1000000"), Plan1.Range("C4:C1000000"), produto)) - (WorksheetFunction.SumIfs(Plan2.Range("d4:d1000000"), Plan2.Range("c4:c1000000"), produto))
End With
linhalistbox = linhalistbox + 1
End If
linha = linha + 1
Wend
End With
End Sub
Sub ordemalfabetica()
Dim i, j, x As Long
Dim temp As String
   With ListBox1
   For j = LBound(.List) To UBound(.List) - 1 Step 1
   For i = LBound(.List) To UBound(.List) - 1 Step 1
   If UCase(.List(i, 0)) > UCase(.List(i + 1, 0)) Then
   
If i <> 0 Then
For x = 0 To (.ColumnCount - 1) Step 1
temp = .List(i, x)
.List(i, x) = .List(i + 1, x)
.List(i + 1, x) = temp
Next x
End If
End If
Next i
Next j
End With

   
   
   
End Sub



Sub propriedadeslb1()
With ListBox1
.Clear
.ColumnWidths = "330;80"
.ColumnCount = 2
.ListStyle = fmListStylePlain
.AddItem
.List(0, 0) = "NOME"
.List(0, 1) = "QUANTIDADE"

End With

End Sub















