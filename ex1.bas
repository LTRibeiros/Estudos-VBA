Attribute VB_Name = "ex1"
Option Base 1
Sub ex1()

Worksheets("ex1").Activate

Dim i As Integer
Dim j As Integer
Dim matriz(7, 2) As Double

For i = 2 To 6
 For j = 1 To 2
 matriz(i, j) = Fvendas(Cells(i, "B"), Cells(i, "C"))
 Cells(i, "D") = matriz(i, j)
 
 If Cells(i, "D") < 1500 Then
 Cells(i, "E") = "Melhorar marketing"
 Else
 Cells(i, "E") = "Está Bom"
 End If
 

Next j
Next i

End Sub

Function Fvendas(B As Double, C As Double) As Double

Dim x As Double
Dim soma As Double
Dim i As Integer

Fvendas = B * C
Cells(7, "C") = "Total de vendas"
Cells(1, "D") = "Total vendas unitárias(R$)"
Cells(1, "E") = "Ação"

For i = 2 To 6
soma = Cells(i, "D") + soma
Cells(7, "D") = soma
Next i

End Function
