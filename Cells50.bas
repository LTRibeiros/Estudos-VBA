Attribute VB_Name = "Módulo2"
Sub vetorteste()

Dim vetor(50) As Integer
Dim soma As Integer
Dim i As Integer
Dim j As Integer


For i = 1 To 50
soma = i
vetor(i) = soma
j = j + 1
Cells(1, j) = vetor(i)
Next i
End Sub
