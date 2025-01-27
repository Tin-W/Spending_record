Attribute VB_Name = "Module4"
'

'Common VarType Constants:
'0 = vbEmpty (Empty)
'1 = vbNull (Null)
'2 = vbInteger (Integer)
'3 = vbLong (Long)
'4 = vbSingle (Single)
'5 = vbDouble (Double)
'6 = vbCurrency (Currency)
'7 = vbDate (Date)
'8 = vbString (String)
'9 = vbObject (Object)
'11 = vbBoolean (Boolean)
'12 = vbVariant (Variant)
'13 = vbDataObject (Data Object)


'DATA STORAGE


Sub name_sheet()
Dim x As Integer
For x = 2 To 7
Sheets(x).Select
Range("A1").Value = "data storage " & ActiveSheet.name
Call rearrange_cell_size
Next x

End Sub



Sub rearrange_cell_size()
'
' rearrange_cell_size ⅷ떠
'

'
    Cells.Select
    Range("C3").Activate
    Cells.EntireColumn.AutoFit
End Sub

Sub BackUp()
'
' BackUp2 ⅷ떠
'

'
Dim n As Integer
n = Worksheets.Count
    Sheets("Record").Select
    Sheets("Record").Copy After:=Sheets(n)
    Sheets("Record (2)").Select
    Sheets("Record (2)").name = "Backup" & n
End Sub


Sub BackUp2()
'
' BackUp2 ⅷ떠
'

'
Dim week As Integer
week = Range("C2").Value
Dim Months As Integer
Months = Range("C6").Value
Dim n As Integer
n = Worksheets.Count

    Sheets("Record").Select
    Sheets("Record").Copy After:=Sheets(n)
    Sheets("Record (2)").Select
    Sheets("Record (2)").name = "Backup month," & Months & "week" & week
    Sheets(1).Select
End Sub

Sub Last_part_of_end_week()
'
' part_of_end_week ⅷ떠
'

'
Dim week As Integer
week = Range("C2").Value
    Range("D2:I2").Select
    Selection.ClearContents
    Range("U14:U15").Select
    Selection.ClearContents
    Range("AC3:AC4").Select
    Selection.ClearContents
    Range("O11:O14").Select
    Selection.ClearContents

    Range("Z4").Select
    Selection.ClearContents
    If week / 4 = 1 Then
            Range("C6").Value = Range("C6").Value + 1
            Range("C2").Value = 1
            Call dtat_transform_bill
            
    Else: Range("C2").Value = Range("C2").Value + 1
    End If


End Sub

Sub search_object()
'
' search_object ⅷ떠
'

'
    
    Cells.Find(What:="spending", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
End Sub


Sub end_of_week()
Call BackUp2

Dim Months As Integer
Months = Range("C6").Value
Dim fee As Integer
fee = 1
Dim userResponse As Integer
    
' Ask the user with a Yes/No prompt
userResponse = MsgBox("Are you sure this is the end of the week?", vbYesNo + vbQuestion, "End of Week Confirmation")
If userResponse = vbNo Then
    Exit Sub
ElseIf userResponse = vbYes Then
    
    Call data_transfer_Society_week
    
    Sheets("Record").Activate
    Cells.Find(What:="spending", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate
    Do Until IsEmpty(ActiveCell.Offset(fee, 0))
    Dim cost As Integer
    cost = Range("C2").Offset(0, fee).Value
    ActiveCell.Offset(fee, Months) = ActiveCell.Offset(fee, Months) + cost
    fee = fee + 1
    Loop
End If
Call dtat_transform_bill
Sheets("Record").Activate
Call Last_part_of_end_week

End Sub

Sub end_of_the_day()
Sheets("Record").Activate

Dim userResponse As Integer
userResponse = MsgBox("have you participate in society and travell by bus/train in London?", vbYesNo + vbQuestion, "Good job man, you made it and completed it well, i am proud of you")
If userResponse = vbNo Then
    MsgBox ("no rush man, take your time and no pressure :)")
    Exit Sub
ElseIf userResponse = vbYes Then

    Call data_transfer_Society_day
    Call data_transfer_entertainment
    Sheets("Record").Activate
    Call Entertainment
    
    Call data_transfer_food
    Sheets("Record").Activate
    Call EatOut
    Call Gorcery
    
    Call data_transform_shopping
    Sheets("Record").Activate
    Call shoping
    
    Call Bike
    
    Call data_transfer_transport
    Sheets("Record").Activate
    Call other_city
    Sheets("Record").Activate
    Range("AC3:AC4") = 0
    Range("V4:Y4").Select
    Selection.ClearContents
    Range("O2").Select
    Selection.ClearContents
    Range("Z3") = 0
End If

End Sub
