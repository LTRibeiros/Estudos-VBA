Attribute VB_Name = "entregas"
Sub encomendas()

Dim N_enc As Integer
Dim T_est As Double
Dim T_med As Integer
Dim Q_enc As Integer
Dim i As Double
Dim p As Double
Dim T_med2 As Double

p = 1
Q_enc = InputBox("Quantos diferentes produtos serão entregues?")
i = Q_enc
Do While i > 0
N_enc = InputBox("Número de encomendas do produto" & " " & p & " " & "entregues em um dia:")

T_est = InputBox("Tempo estimado de cada entrega do produto" & " " & p & " " & "(em dias):")

T_med = N_enc / T_est

MsgBox ("O tempo médio de entrega do produto" & " " & p & " " & "será de:" & " " & T_med & " " & "dias")
i = i - 1
p = p + 1
'T_med = N_enc / T_est
'MsgBox ("O tempo médio de entrega será de:" & " " & T_med & " " & "dias")
Loop

MsgBox ("o menor tempo de entrega será o do produto" & T_med & "e o maior tempo de entrega será do produto" & T_med2)



End Sub
