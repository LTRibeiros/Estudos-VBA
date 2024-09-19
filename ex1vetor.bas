Attribute VB_Name = "Módulo1"
Option Explicit

Sub vetorex1()
Worksheets("vetor50").Activate
Dim vetor(50) As Integer
Dim i As Integer
Dim den As Integer
Dim vet_soma As Integer
Dim media As Double
Dim j As Integer
Dim impar As Integer

For i = 1 To 50
j = j + 1
vetor(i) = Cells(1, j)
If vetor(i) > 10 And vetor(i) < 200 Then
den = den + 1
vet_soma = vet_soma + vetor(i)
End If
Next i
media = vet_soma / den
MsgBox (media)

j = 0
For i = 1 To 50
j = j + 1
vetor(i) = Cells(1, j)
If vetor(i) Mod 2 = 1 Then
impar = vetor(i) + impar

End If
Next i
MsgBox (impar)

End Sub
