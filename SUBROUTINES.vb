

Sub enter()

Dim rang As Range
Dim rang2 As Range

'RANGE FOR ACCOUNT CODES
Set rang = Range("b10", Range("b10").End(xlDown))
Set rang2 = Range("c5:c6")

    If rang.Rows.Count > 200 Or rang.Rows.Count < 2 Then
    MsgBox ("Check if you have skipped a line, the Count of rows is too big or too small")
    Exit Sub
    End If

Dim Sure As Integer
 Sure = MsgBox("Are you sure?", vbOKCancel)
 If Sure = 2 Then
      Exit Sub
 End If


Dim aStrings(0 To 10) As String
aStrings(1) = "348": aStrings(2) = "378": aStrings(2) = "321":  aStrings(3) = "391":
aStrings(3) = "503": aStrings(4) = "515": aStrings(5) = "519":  aStrings(6) = "705":
aStrings(7) = "709": aStrings(8) = "711": aStrings(9) = "713": aStrings(10) = "378":

For Each i In aStrings
On Error Resume Next
Range("b10:b199").Select
code = Range("b10:b199").Find(What:=i, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
pvalue = Range("c" & code).Value
pvaluea = Range("c" & code)
If pvalue > 0 Then
Range("c" & code).Value = -1 * Range("c" & code).Value
End If
Next i



    If IsEmpty(Range("B10").Value) = True Then
      MsgBox "The First Line is Empty is empty, Please Double Check"
      Exit Sub
    End If
     If IsEmpty(Range("B10").Value) = True Then
      MsgBox "The First Line is Empty is empty, Please Double Check"
      Exit Sub
    End If
    If IsEmpty(Range("C5").Value) = True Then
      MsgBox "Balance is empty, Please Double Check"
      Exit Sub
    End If
    If IsEmpty(Range("c6").Value) = True Then
      MsgBox "Current Year is empty, Please Double Check"
      Exit Sub
    End If
    If IsEmpty(Range("c8").Value) = True Then
      MsgBox "Type of Form is empty, Please Double Check"
      Exit Sub
    End If

    If Round(Range("E5").Value, 0) <> Round(Range("E25").Value, 0) Then
      MsgBox "Activos Corrientes dont match (339), the difference is " & Round(Range("E5").Value, 0) - Round(Range("E25").Value, 0)
      Exit Sub
    End If
     If Round(Range("E6").Value, 0) <> Round(Range("E26").Value, 0) Then
      MsgBox "Activos Fijos dont match  (369) the difference is " & Round(Range("E6").Value, 0) - Round(Range("E26").Value, 0)
      Exit Sub
    End If
     If Round(Range("E8").Value, 0) <> Round(Range("E27").Value, 0) Then
      MsgBox "Activos Diferidos Dont Match  (379) the difference is" & Round(Range("E8").Value, 0) - Round(Range("E27").Value, 0)
      Exit Sub
    End If
     If Round(Range("E9").Value, 0) <> Round(Range("E28").Value, 0) Then
      MsgBox "Activos Largo Plazo Dont Match (397) the difference is" & Round(Range("E9").Value, 0) - Round(Range("E28").Value, 0)
      Exit Sub
    End If
     If Round(Range("E10").Value, 0) <> Round(Range("E29").Value, 0) Then
      MsgBox "Total Activos Dont Match (399) the difference is" & Round(Range("E10").Value, 0) - Round(Range("E29").Value, 0)
      Exit Sub
    End If

    If Round(Range("G5").Value, 0) <> Round(Range("E30").Value, 0) Then
      MsgBox "Total-Pasivo Corriente(439) Dont Match the difference is" & Round(Range("G5").Value, 0) - Round(Range("E30").Value, 0)
      Exit Sub
    End If
     If Round(Range("G6").Value, 0) <> Round(Range("E31").Value, 0) Then
      MsgBox "Total Pasivo Largo Plazo (469) Dont Match the difference is" & Round(Range("G6").Value, 0) - Round(Range("E31").Value, 0)
      Exit Sub
    End If

    If Round(Range("G9").Value, 0) <> Round(Range("E32").Value, 0) Then
      MsgBox "Total Pasivo (499) Dont Match the difference is" & Round(Range("G9").Value, 0) - Round(Range("E32").Value, 0)
      Exit Sub
    End If
     If Round(Range("G10").Value, 0) <> Round(Range("E33").Value, 0) Then
      MsgBox "Total Ingresos (699) Dont Match the difference is" & Round(Range("G10").Value, 0) - Round(Range("E33").Value, 0)
      Exit Sub
    End If
     If Round(Range("G11").Value, 0) <> Round(Range("E34").Value, 0) Then
      MsgBox "Total Costos y Gastos (799) Dont Match the difference is" & Round(Range("G11").Value, 0) - Round(Range("E34").Value, 0)
      Exit Sub
    End If

 Call VBA_to_append_existing_text_file

End Sub


Sub VBA_to_append_existing_text_file()
    Dim allowed As Integer
    allowed = 0
    Call check(0, allowed) 'CONDITIONAL STATEMENTS FOR CHECKS, IF THE BALANCE IS NOT IN SAMPLE, WONT ALLOW TO MOVE
    If allowed <> 1 Then
        MsgBox ("Function is not in SAMPLE or already entered")
        Exit Sub
    End If
'SETTING UP VARIABLES
    Dim expediente, fechas As Integer
    Dim strFile_Path As String
    Dim rang, rang2 As Range
     
'RANGE FOR ACCOUNT CODES
    Set rang = Range("b10", Range("b10").End(xlDown))
    Set rang2 = Range("c5:c6")

'This sets up the directory
    strFile_Path = ThisWorkbook.Path & "\OUTPUT\Balances.txt"
    Open strFile_Path For Append As #1
    For Each element In rang
    Dim a As Integer
    a = element.Row
    Print #1, Range("C5").Value & "@" & Range("c6").Value & "@" & Range("c8").Value & "@" & Range("B" & a) & "@" & Range("C" & a)
    Next element
    Close #1
    
    strFile_Pathb = ThisWorkbook.Path & "\OUTPUT\INDIVIDUAL\"
    Dim file_output As String
    file_output = Range("C5").Value & "_" & Range("c6").Value
    strFile_Pathb = strFile_Pathb + "Bal_" + file_output + ".txt"
    Open strFile_Pathb For Append As #1
    For Each element In rang
    Dim b As Integer
    b = element.Row
    Print #1, Range("C5").Value & "@" & Range("c6").Value & "@" & Range("c8").Value & "@" & Range("B" & b) & "@" & Range("C" & b)
    Next element
    Close #1
'Writing a One on the ones that is done
    number = ThisWorkbook.Sheets("BALANCES").Range("i1:i1")
    ThisWorkbook.Sheets("BALANCES").Range("D" & number).Value = 1
    Set ranga = Range("b10", Range("b10").End(xlDown))
    Set rangb = Range("c10", Range("c10").End(xlDown))
    Set rangc = Range("c5:c8")
    ranga.ClearContents
    rangb.ClearContents
    rangc.ClearContents
    MsgBox ("Submitted")
End Sub

Sub clear()

     Dim rang As Range
     Dim rang2 As Range
'Looking at the Range
    Set rang1 = Range("b10", Range("b10").End(xlDown))
    Set rang2 = Range("c10", Range("c10").End(xlDown))
    Set rang3 = Range("c5:c8")
    Dim Sure As Integer
    Sure = MsgBox("Are you sure You want to Clear?", vbOKCancel)
    If Sure = 2 Then
         Exit Sub
    End If
    rang1.ClearContents
    rang2.ClearContents
    rang3.ClearContents
   
End Sub

Sub getnextbalance()

     Dim rang As Range
     Dim number As Integer
     Dim expran As String
     Dim expran2 As Range
     Dim numbera As Integer
     Dim rang2 As Range
'Looking at the Range
     Set rang = Range("b10", Range("b10").End(xlDown))
     Set rang3 = Range("C10", Range("C10").End(xlDown))
     Set rang2 = Range("c5:c6")
     rang.ClearContents
     rang2.ClearContents
     rang3.ClearContents
     ThisWorkbook.Sheets("BALANCES").Range("c1").FormulaR1C1 = "=MATCH(0,R[1]C:R[5000]C,0)"
     number = ThisWorkbook.Sheets("BALANCES").Range("D1:D1")
     ThisWorkbook.Sheets("BALANCES").Range("B" & number + 1).Copy
     ThisWorkbook.Sheets("FORM").Range("C5").PasteSpecial Paste:=xlPasteValues
     ThisWorkbook.Sheets("BALANCES").Range("C" & number + 1).Copy
     ThisWorkbook.Sheets("FORM").Range("C6").PasteSpecial Paste:=xlPasteValues
     
   
End Sub

Public Function check(tipo As Integer, allowed As Integer)
    Dim code As String
    Match = ThisWorkbook.Sheets("BALANCES").Range("I1:I1")
    If IsError(Match) Then
    MsgBox ("NOT IN SAMPLE- PLEASE CHECK BALANCE NUMBER")
    Exit Function
    End If
    
    code = ThisWorkbook.Sheets("BALANCES").Range("h1:h1")
    If IsError(code) Then
    MsgBox ("NOT IN SAMPLE- PLEASE CHECK BALANCE NUMBER")
    Exit Function
    End If

    
    If tipo = 1 And code = 1 Then
        MsgBox ("ALREADY ENTERED")
        Exit Function
    ElseIf tipo = 1 And code = 0 Then
        MsgBox ("NOT YET ENTERED")
        Exit Function
    ElseIf tipo = 0 And code = 0 Then
        allowed = 1
        End If

End Function

Private Sub CommandButton22_Click()
Call clear
End Sub

Private Sub CommandButton21_Click()
Call enter
End Sub

Private Sub CommandButton23_Click()
Dim allowed As Integer
allowed = 0
Call check(1, allowed)
End Sub


Private Sub CommandButton24_Click()
Call getnextbalance
End Sub



