Attribute VB_Name = "exvetor8"
'8.  Criar e carregar uma matriz [4][3] inteiro com quantidade de produtos vendidos em 4 semanas. Calcular e exibir:
'a.  A quantidade de cada produto vendido no mês;
'b.  A quantidade de produtos vendidos por semana;
'c.  O total de produtos vendidos no mês.
Option Base 1

Sub ex8matriz()
Worksheets("produtos").Activate




Dim matriz(4, 3) As Integer
Dim c As Integer
Dim i As Integer
Dim j As Integer
Dim produtos_semana As Integer
Dim x As Integer



j = 1
For i = 1 To 4
 For j = 1 To 3
  ii = i + 1
 jj = j + 1
 
 matriz(i, j) = InputBox(i & " x " & j)
 Cells(ii, jj) = matriz(i, j)
 somaprodutos = somaprodutos + matriz(i, j)

 
 Next j
 Next i
 MsgBox ("o total de produtos vendidos no mês foi de: " & somaprodutos)
For i = 1 To 4
x = x + 1
 For j = 1 To 3
 produtos_semana = produtos_semana + matriz(x, j)
 
 Next j
 MsgBox "produtos vendidos na semana " & i & " " & produtos_semana
 produtos_semana = 0
 Next i
 
For j = 1 To 3
produto_mes = 0

  For i = 1 To 4

  ii = i + 1
 jj = j + 1
 
 produto_mes = produto_mes + Cells(ii, jj)
 
Next i
MsgBox ("O produto " & j & " foi vendido " & produto_mes & " vezes no mês")

Next j


End Sub


 
 
