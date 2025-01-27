Attribute VB_Name = "Module2"
Sub EatOut()
'
'
Range("E2").Value = Range("E2").Value + Range("N7").Value
Range("N7").Value = 0
End Sub

Sub Gorcery()
'
'
Range("E2").Value = Range("E2").Value + Range("O7").Value
Range("O7").Value = 0
End Sub


Sub SocietyA()
'
' Society ⅷ떠
Range("O11").Value = Range("O11").Value + 1
Range("G2").Value = Range("G2").Value + Range("N11").Value + Range("P11")
Range("P11").Value = 0
End Sub

Sub SocietyB()
'
' Society ⅷ떠
Range("O12").Value = Range("O12").Value + 1
Range("G2").Value = Range("G2").Value + Range("N12").Value + Range("P12")
Range("P12").Value = 0
End Sub

Sub SocietyC()
'
' Society ⅷ떠
Range("O13").Value = Range("O13").Value + 1
Range("G2").Value = Range("G2").Value + Range("N13").Value + Range("P13")
Range("P13").Value = 0
End Sub

Sub SocietyD()
'
' Society ⅷ떠
Range("O14").Value = Range("O14").Value + 1
Range("G2").Value = Range("G2").Value + Range("N14").Value + Range("P14")
Range("P14").Value = 0
End Sub

Sub shoping()
    Dim x As Integer
    Dim spend As Integer

    For x = 0 To 3
        ' Check if the value is not a Double (type 5) and is not equal to 0
        If VarType(Range("T" & 7 + x).Value) = 0 Or 5 Then
            ' Get the value from the T column and add it to the spend
            spend = Cells(7 + x, 20).Value
            Cells(2, 8).Value = Cells(2, 8).Value + spend
            ' Reset the value in the T column to 0
            Cells(7 + x, 20).Value = 0
        Else
            MsgBox ("Input is in incorrect form")
            Range("T" & 7 + x).Select
            Exit Sub
        End If
    Next x
End Sub

Sub Entertainment()
'
' Entertainment ⅷ떠
'

'
    Dim options As Integer
    For options = 1 To 3
    
    If VarType(Range("T" & 13 + options).Value) = 5 Then

        If Range("T" & 13 + options).Value <> 0 Then
            Range("T" & 13 + options).Offset(0, 1).Value = Range("T" & 13 + options).Offset(0, 1).Value + 1
            Range("F2").Value = Range("F2").Value + Range("T" & 13 + options).Value
            Range("T" & 13 + options) = 0
        Else
            Range("T" & 13 + options) = 0
        End If
    ElseIf VarType(Range("T" & 13 + options).Value) = 0 Then
        Range("T" & 13 + options) = 0
    Else
        MsgBox ("input is incorrect form")
        Range("T" & 13 + options).Select
        Exit Sub
    End If
    Next options
End Sub


