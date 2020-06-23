VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Relatório 
   Caption         =   "Relatório Financeiro"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   OleObjectBlob   =   "Form_Relatório.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Relatório"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adicionar_Click()
Call financeiro
End Sub
Sub financeiro()
If Data1.Value <> "" And Data2.Value <> "" Then
On Error GoTo erro
Datainicial = CDbl(DateValue(Data1))
Datafinal = CDbl(DateValue(Data2))
Dim valorentrada As Double
Dim valorsaida As Double
Dim valorganhos As Double
Dim valorgastos As Double
Dim total As Double
valorentrada = WorksheetFunction.SumIfs(Plan1.Range("f4:f1000000"), Plan1.Range("g4:g1000000"), ">=" & Datainicial, Plan1.Range("g4:g1000000"), "<=" & Datafinal)
valorsaida = WorksheetFunction.SumIfs(Plan2.Range("f4:f1000000"), Plan2.Range("g4:g1000000"), ">=" & Datainicial, Plan2.Range("g4:g1000000"), "<=" & Datafinal)
valorganhos = WorksheetFunction.SumIfs(Plan4.Range("b4:b1000000"), Plan4.Range("d4:d1000000"), ">=" & Datainicial, Plan4.Range("d4:d1000000"), "<=" & Datafinal)
valorgastos = WorksheetFunction.SumIfs(Plan5.Range("b4:b1000000"), Plan5.Range("d4:d1000000"), ">=" & Datainicial, Plan5.Range("d4:d1000000"), "<=" & Datafinal)
total = CDbl(valorsaida) - CDbl(valorentrada) - CDbl(valorgastos) + CDbl(valorganhos)
Label1.Caption = "Você comprou " & "R$ " & FormatNumber(valorentrada, 2) & " entre " & Data1.Text & " e " & Data2.Text
Label2.Caption = "Você vendeu " & "R$ " & FormatNumber(valorsaida, 2) & " entre " & Data1.Text & " e " & Data2.Text
Label3.Caption = "Você gastou " & "R$ " & FormatNumber(valorgastos, 2) & " entre " & Data1.Text & " e " & Data2.Text
Label4.Caption = "Seus ganhos extras foram " & "R$ " & FormatNumber(valorganhos, 2) & " entre " & Data1.Text & " e " & Data2.Text
Label5.Caption = "O saldo total entre " & Data1.Text & " e " & Data2.Text & " foi de: " & "R$ " & FormatNumber(total, 2)
Label7.Caption = FormatNumber(total, 2)
Else
MsgBox "Preencha todos os campos", vbInformation, "Erro"
End If
Exit Sub
erro:
MsgBox "Digite uma data válida", vbInformation, "Erro"
Exit Sub
End Sub
