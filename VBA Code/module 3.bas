Attribute VB_Name = "Module3"



Sub data_transfer_transport()
'creat a section in each worksheet, transfer a total amount spended and number of each specific event
'for shopping, need to tranfer every time we click "shopping"
Sheets("transport").Activate

Dim target As Long
target = Cells(Rows.Count, 1).End(xlUp).ROW + 1
Cells(target, 1).Value = Date
Cells(target, 1).Offset(2, 0).Value = "Number"
Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate

Dim names As Variant
    names = Array("", "BUS:", "Zone 1", "Zone 2", "Zone 3", "Zone 4", "Other city", "Bike:")

Dim subtitle As Integer

For subtitle = 1 To UBound(names)
   Dim s_t As Range
   ActiveCell.Offset(0, subtitle) = names(subtitle)

   If ActiveCell.Offset(0, subtitle).Value = "BUS:" Then
        With Sheets("Record").Cells
        Set s_t = .Find(What:=names(subtitle), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        ActiveCell.Offset(1, subtitle).Value = s_t.Offset(0, 1).Value * s_t.Offset(0, 2).Value
        End With
    ElseIf ActiveCell.Offset(0, subtitle).Value = "Zone 1" Or ActiveCell.Offset(0, subtitle).Value = "Zone 2" Or ActiveCell.Offset(0, subtitle).Value = "Zone 3" Or ActiveCell.Offset(0, subtitle).Value = "Zone 4" Then
        With Sheets("Record").Cells
        Set s_t = .Find(What:=names(subtitle), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        ActiveCell.Offset(1, subtitle) = s_t.Offset(1, 0).Value * s_t.Offset(2, 0).Value
        End With

    ElseIf ActiveCell.Offset(0, subtitle).Value = "Other city" Then
        With Sheets("Record").Cells
        Set s_t = .Find(What:=names(subtitle), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        ActiveCell.Offset(1, subtitle) = s_t.Offset(1, 0).Value
        End With
    ElseIf ActiveCell.Offset(0, subtitle).Value = "Bike:" Then
        'Range("D2").Value + Range("AC3") * Range("AF3") + Range("AC4") * Range("AF4")
        With Sheets("Record").Cells
        Set s_t = .Find(What:=names(subtitle), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        ActiveCell.Offset(1, subtitle) = s_t.Offset(1, 1).Value * s_t.Offset(1, 4) + s_t.Offset(2, 1).Value * s_t.Offset(2, 4)
        End With

    Else
        MsgBox ("VBA error")
        Exit Sub
    End If
Next subtitle


Cells.Find(What:="Number", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate
'n_l=number loop
Dim n_l As Integer
For n_l = 1 To UBound(names)
    Dim s_t2 As Range

    If ActiveCell.Offset(-2, n_l).Value = "BUS:" Then
            With Sheets("Record").Cells
            Set s_t2 = .Find(What:=names(n_l), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            ActiveCell.Offset(0, n_l).Value = s_t2.Offset(0, 2).Value
            End With
    ElseIf ActiveCell.Offset(-2, n_l).Value = "Zone 1" Or ActiveCell.Offset(-2, n_l).Value = "Zone 2" Or ActiveCell.Offset(-2, n_l).Value = "Zone 3" Or ActiveCell.Offset(-2, n_l).Value = "Zone 4" Then
           With Sheets("Record").Cells
           Set s_t2 = .Find(What:=names(n_l), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
           If s_t2.Offset(2, 0).Value = vbEmpty Then
                ActiveCell.Offset(0, n_l) = 0
           Else
                ActiveCell.Offset(0, n_l) = s_t2.Offset(2, 0)
           End If
           End With
    
    ElseIf ActiveCell.Offset(-2, n_l).Value = "Other city" Then
            If ActiveCell.Offset(-1, n_l).Value = 0 Then
               ActiveCell.Offset(0, n_l).Value = "not apply"
            Else
                ActiveCell.Offset(0, n_l).Value = InputBox("where did you go?")
    
            End If

    ElseIf ActiveCell.Offset(-2, n_l).Value = "Bike:" Then
           With Sheets("Record").Cells
           Set s_t2 = .Find(What:=names(n_l), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
           'If Range("AC3") = 0 Or Range("AC3").Value = "0" Then
            ActiveCell.Offset(0, n_l) = s_t2.Offset(1, 1).Value
           ActiveCell.Offset(0, n_l + 1).Value = "Munite:" & s_t2.Offset(2, 1).Value
           End With
                  
    Else
        MsgBox "VBA error", vbOKCancel, "Error"
        Exit Sub
    End If
Next n_l

End Sub

Sub data_transfer_food()
Sheets("food").Activate
Dim role As Long
role = Cells(Rows.Count, 1).End(xlUp).ROW + 1
Cells(role, 1).Value = Date
Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate

Dim num As Integer
num = 1
Do Until num = 3
    'tp = spending type
    Dim tp As Range
    With Sheets("Record").Cells
    Set tp = .Find(What:="Food:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    ActiveCell.Offset(0, num).Value = tp.Offset(1, num).Value
    End With
    num = num + 1
Loop
End Sub

Sub dtat_transform_bill()

Sheets("bills").Activate
Dim role As Long
role = Cells(Rows.Count, 1).End(xlUp).ROW + 1
Cells(role, 1).Value = Date
Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate
Dim hi As Range
With Sheets("Record").Cells
Set hi = .Find(What:="Bill", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

ActiveCell.Offset(0, 1).Value = hi.Offset(1, 0).Value
End With

End Sub

Sub data_transfer_entertainment()

Sheets("entertainment").Activate
Dim role As Long
role = Cells(Rows.Count, 1).End(xlUp).ROW + 1
Cells(role, 1).Value = Date
Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, MatchByte:=False, SearchFormat:=False).Activate
ActiveCell.Offset(1, 0) = "cost:"
Dim name As Variant
    name = Array("", "clubbing", "party", "other")

Dim subt As Integer
subt = 1
Do Until subt = UBound(name) + 1
   Dim st As Range
   ActiveCell.Offset(0, subt) = name(subt)
   If ActiveCell.Offset(0, subt).Value = "clubbing" Or ActiveCell.Offset(0, subt).Value = "party" Then
        With Sheets("Record").Cells
        Set st = .Find(What:=name(subt), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        ActiveCell.Offset(1, subt).Value = st.Offset(0, 1).Value
        End With
        
   ElseIf ActiveCell.Offset(0, subt).Value = "other" Then
        
        ActiveCell.Offset(1, subt).Value = Sheets("Record").Range("T16").Value
        If ActiveCell.Offset(1, subt).Value = 0 Then
            ActiveCell.Offset(1, subt + 1).Value = "good"
        Else
            ActiveCell.Offset(1, subt + 1).Value = InputBox("what did you do for other entertainment?")
        End If
   Else
       MsgBox "error", vbOKOnly, "VBA Background module 3"
   End If
   subt = subt + 1
Loop
    
End Sub


Sub data_transform_shopping()
    Sheets("shopping").Activate
    
    Dim ROW As Long
    ROW = Cells(Rows.Count, 1).End(xlUp).ROW + 1
    Cells(ROW, 1).Value = Date

    Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt:=xlPart, _
               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, _
               MatchByte:=False, SearchFormat:=False).Activate
    
    ActiveCell.Offset(1, 0) = "cost:"
    
    Dim list As Variant
    list = Array("", "clothes", "shoes", "Luxury", "needs")
    
    Dim su As Integer
    su = 1
    
    Do Until su = UBound(list) + 1
        Dim s As Range
        ActiveCell.Offset(0, su) = list(su)
        
        If ActiveCell.Offset(0, su).Value = "clothes" Or ActiveCell.Offset(0, su).Value = "shoes" Or ActiveCell.Offset(0, su).Value = "Luxury" Or ActiveCell.Offset(0, su).Value = "needs" Then
            With Sheets("Record").Cells
                Set s = .Find(What:=list(su), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                If Not s Is Nothing Then
                    ActiveCell.Offset(1, su).Value = s.Offset(0, 1).Value
                Else
                    MsgBox "Value not found in 'Record' sheet.", vbOKOnly, "Error"
                End If
            End With
        End If
        
        su = su + 1
    Loop
    
End Sub

Sub data_transfer_Society_day()
Sheets("society").Activate
Dim ROW As Long
    ROW = Cells(Rows.Count, 1).End(xlUp).ROW + 1
    Cells(ROW, 1).Value = Date

    Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt:=xlPart, _
               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, _
               MatchByte:=False, SearchFormat:=False).Activate
    
    ActiveCell.Offset(1, 0) = "extra cost:"
    Dim name_list As Variant
    Dim A As Variant
    A = Sheets("Record").Range("M11").Value
    Dim B As Variant
    B = Sheets("Record").Range("M12").Value
    Dim C As Variant
    C = Sheets("Record").Range("M13").Value
    Dim D As Variant
    D = Sheets("Record").Range("M14").Value
    
    name_list = Array("", A, B, C, D)
    
    Dim diu As Integer
    diu = 1
    
    Do Until diu = UBound(name_list) + 1
        Dim fuck As Range
        ActiveCell.Offset(0, diu) = name_list(diu)
        
        If ActiveCell.Offset(0, diu).Value = A Or ActiveCell.Offset(0, diu).Value = B Or ActiveCell.Offset(0, diu).Value = C Or ActiveCell.Offset(0, diu).Value = D Then
            With Sheets("Record").Cells
                Set fuck = .Find(What:="Extra Event", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                If Not fuck Is Nothing Then
                    ActiveCell.Offset(1, diu).Value = fuck.Offset(diu, 0).Value
                Else
                    MsgBox "Value not found in 'Record' sheet.", vbOKOnly, "Error"
                End If
            End With
        End If
        
        diu = diu + 1
    Loop
End Sub

Sub data_transfer_Society_week()

Sheets("society").Activate
Dim ROW As Long
    ROW = Cells(Rows.Count, 1).End(xlUp).ROW + 1
    Cells(ROW, 1).Value = Date

    Cells.Find(What:=Date, After:=ActiveCell, LookIn:=xlFormulas2, LookAt:=xlPart, _
               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, _
               MatchByte:=False, SearchFormat:=False).Activate
    
    ActiveCell.Offset(1, 0) = "week cost:"
    ActiveCell.Offset(2, 0) = "week number" & Sheets("Record").Range("C2").Value
    Dim name_list As Variant
    Dim A As Variant
    A = Sheets("Record").Range("M11").Value
    Dim B As Variant
    B = Sheets("Record").Range("M12").Value
    Dim C As Variant
    C = Sheets("Record").Range("M13").Value
    Dim D As Variant
    D = Sheets("Record").Range("M14").Value
    
    name_list = Array("", A, B, C, D)
    
    Dim damn As Integer
    damn = 1
    
    Do Until damn = UBound(name_list) + 1
        Dim shit As Range
        ActiveCell.Offset(0, damn) = name_list(damn)
        
        If ActiveCell.Offset(0, damn).Value = A Or ActiveCell.Offset(0, damn).Value = B Or ActiveCell.Offset(0, damn).Value = C Or ActiveCell.Offset(0, damn).Value = D Then
            With Sheets("Record").Cells
                
                Set shit = .Find(What:="CPT", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                
                If IsNumeric(shit.Offset(damn, 0).Value) And IsNumeric(shit.Offset(damn, 1).Value) Then
                ActiveCell.Offset(1, damn).Value = shit.Offset(damn, 0).Value * shit.Offset(damn, 1).Value
                 
                Else
                    MsgBox "some value is not number", vbOKOnly, "Error"
                    Sheets("Record").Activate
                    shit.Offset(damn, 0).Select
                    Exit Sub
                    
                End If
            End With
        End If
        
        damn = damn + 1
    Loop
End Sub

