Attribute VB_Name = "Módulo1"
Sub ex2vetor()

Dim vetor(100) As Integer
Dim media As Double
Dim soma As Integer
Dim maior As Integer
Dim menor As Integer
Dim j As Integer
Dim den As Integer
maior = vetor(1)
menor = vetor(1)
For i = 1 To 100
 j = j + 1
   vetor(i) = Cells(1, j)
If maior < vetor(i) Then
   maior = vetor(i)
End If
If menor > vetor(i) Then
   menor = vetor(i)
End If

soma = soma + vetor(i)
den = den + 1
Next i
media = soma / den
MsgBox (maior)
MsgBox (menor)
MsgBox (media)

End Sub
