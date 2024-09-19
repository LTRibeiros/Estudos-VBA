Attribute VB_Name = "exvetor4"
Option Base 1

Sub ex4vetor()

Worksheets("media").Activate

Dim vetor(30) As Double
Dim j As Integer

For i = 1 To 30
vetor(i) = InputBox("Nota do aluno " & i)
j = j + 1

Cells(j, 1) = ("Aluno " & j)
Cells(j, 2) = vetor(i)


soma = soma + vetor(i)
Next i

media = soma / 30
MsgBox ("A média da turma foi de: " & media)

For i = 1 To 30
If vetor(i) > media Then
cont = cont + 1
Else
MsgBox ("Aluno de número " & i & " (" & ((i & "x" & "2")) & ")" & " teve um desempenho menor que a média do grupo")

End If
Next i

MsgBox ("Quantidade de notas acima da média: " & cont)



End Sub

