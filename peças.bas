Attribute VB_Name = "pe�as"
Sub fabrica()
Dim lote As Integer
Dim Q_ins As Double
Dim x As Double
Dim p As Double
Dim i As Double
Dim n_pi As Double
Dim n_def As Double


lote = InputBox("Qual o tamanho do lote a ser inspecionado?")
Q_ins = lote * 0.1
p = 1
n_pi = 0
i = Q_ins
Do While i > 0
x = InputBox("A pe�a de n�mero" & " " & p & " " & "foi aprovada?" & "                    " & "1 - aprovada 2 - defeituosa")
 If x = 1 Or 2 Then
n_pi = n_pi + 1
 If x = 2 Then
 n_def = n_def + 1
 End If
 End If

i = i - 1
p = p + 1
'MsgBox ("O tempo m�dio de entrega ser� de:" & " " & T_med & " " & "dias")
Loop
MsgBox ("n�mero de pe�as do lote:" & " " & lote)
MsgBox ("n�mero de pe�as inspecionadas:" & " " & n_pi)
MsgBox ("n�mero de pe�as defeituosas:" & " " & n_def)
MsgBox ("porcentagem de pe�as defeituosas em rela��o as pe�as inspecionadas:" & " " & n_def / n_pi * 100 & "%")


End Sub

