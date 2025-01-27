Attribute VB_Name = "Module1"
Sub Bus()
Attribute Bus.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Bus ⅷ떠
Range("O2").Value = Range("O2").Value + 1
Range("D2").Value = Range("D2").Value + Range("N2").Value

End Sub

Sub Zone1()
'
' train ⅷ떠
Range("V4").Value = Range("V4").Value + 1
Range("D2").Value = Range("D2").Value + Range("V3").Value
End Sub

Sub Zone2()
'
' train ⅷ떠
Range("W4").Value = Range("W4").Value + 1
Range("D2").Value = Range("D2").Value + Range("W3").Value
End Sub

Sub Zone3()
'
' train ⅷ떠
Range("X4").Value = Range("X4").Value + 1
Range("D2").Value = Range("D2").Value + Range("X3").Value
End Sub

Sub Zone4()
'
' train ⅷ떠
Range("Y4").Value = Range("Y4").Value + 1
Range("D2").Value = Range("D2").Value + Range("Y3").Value
End Sub

Sub other_city()
'
' train ⅷ떠
Range("D2").Value = Range("D2").Value + Range("Z3").Value
Range("Z3") = 0
End Sub

Sub Bike()
'
'
Range("AC3").Value = InputBox("how many times did u ride a bike today?")


     If Range("AC3") = 0 Or Range("AC3").Value = "0" Then
             Exit Sub
     Else
            Range("AC4").Value = InputBox("how long did u ride a bike today?")
            If VarType(Range("AC3").Value) And VarType(Range("AC4").Value) <> vbDouble Then
                MsgBox "Input is in incorrect form"
                Range("AC3:AC4").Select
                Exit Sub
            Else
                Range("D2").Value = Range("D2").Value + Range("AC3") * Range("AF3") + Range("AC4") * Range("AF4") + Range("AF5")
            End If
     End If
End Sub


