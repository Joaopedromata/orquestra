VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastrar 
   Caption         =   "Produtos Cadastrados"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   OleObjectBlob   =   "Form_Cadastrar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adicionar2_Click()
Nome = UCase(Nome.Value)
Call adicionarnome
Call propriedadeslb1
Call Carregarlv
Call ordemalfabetica
End Sub
Sub Carregarlv()
Dim ultimalinha As Long
Dim linha As Integer
ultimalinha = Plan3.Range("C1000000").End(xlUp).Row
For linha = 4 To ultimalinha
ListBox1.AddItem Plan3.Range("C" & linha)
Next
End Sub
Sub adicionarnome()
Sheets("plan3").Select
Range("c3").Select
Dim msg As String
If Nome.Value = "" Then
msg = MsgBox("Insira um Nome", vbInformation, "Erro")
Else
Do
If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until ActiveCell.Offset(0, 0).Value = Nome.Value Or IsEmpty(ActiveCell) = True
If ActiveCell.Offset(0, 0).Value = Nome.Value Then
msg = MsgBox("Nome Já Existente no Banco de Dados", vbInformation, "Erro")
Else
 ActiveCell.Offset(0, 0).Value = Nome
 End If
 End If
 Nome.Value = ""
 ListBox1.Clear
 
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
Private Sub propriedadeslb1()
With ListBox1
.Clear
.ColumnWidths = "300"
.ColumnCount = 1
.ListStyle = fmListStylePlain
.AddItem
.List(0, 0) = "NOME"
End With
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
End With
linhalistbox = linhalistbox + 1
End If
linha = linha + 1
Wend
End With
End Sub
Private Sub Excluir()
Resp = MsgBox("Deseja Excluir Esses Dados?", vbYesNo, "Excluir")
If Resp = vbYes Then
For i = 0 To (ListBox1.ListCount - 1)
If ListBox1.Selected(i) = True Then
Dim Nome As String
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
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim i As Integer
If Entrada.Visible = True Then
i = ListBox1.ListIndex
Entrada.Nome1.Value = ListBox1.List(i, 0)
Unload Me
ElseIf Saída.Visible = True Then
i = ListBox1.ListIndex
Saída.Nome1.Value = ListBox1.List(i, 0)
Unload Me
End If
End Sub
Private Sub Pesquisar_Change()
Call PesquisarLV
End Sub
Private Sub Remover_Click()
Call Excluir
End Sub
Private Sub UserForm_Initialize()
Call propriedadeslb1
Call Carregarlv
Call ordemalfabetica
End Sub







