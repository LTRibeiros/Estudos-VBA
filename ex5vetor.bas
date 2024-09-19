Attribute VB_Name = "Módulo3"
Option Base 1

Sub ex5vetor()

Dim vetor(20) As Integer
Dim i As Integer
Dim j As Integer

For i = 1 To 20
vetor(i) = InputBox("Valor " & i)
Next i

j = 21


For i = 1 To 10
j = j - 1
x = vetor(i) - vetor(j)
MsgBox (vetor(i) & " - " & vetor(j) & " = " & x)
Next i



End Sub
