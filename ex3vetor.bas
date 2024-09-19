Attribute VB_Name = "Módulo1"
Sub ex3vetor()

Dim VT1(3) As Integer
Dim VT2(3) As Integer
Dim VT3(6) As Integer
Dim j As Integer
Dim y As Integer



For i = 1 To 3
VT1(i) = i
y = y + 1
Cells(1, y) = VT1(i)
Next i

y = 0
For i = 1 To 3
VT2(i) = i + 3
y = y + 1
Cells(2, y) = VT2(i)
Next i
j = 3
y = 0
For i = 1 To 3
VT3(i) = VT1(i)
c = 3
c = c + 1

j = j + 1
y = y + 1
VT3(c) = VT2(i)

Cells(3, y) = VT3(i)
Cells(3, j) = VT3(c)
Next i
End Sub
