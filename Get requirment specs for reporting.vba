Option Explicit
'1 Copy factors and english_short
Sub Copyfactors()
    Dim tab7 As String
    Dim cell As Range
    Dim fname As String
    Dim i As Integer
    Dim tabname As String
    Dim sht As Worksheet
    Dim celladd As String
    Dim copycell As Range
    
    tab7 = "7 FactorsFunctionsPriorities"
    Worksheets(tab7).Activate
   
    fname = "Factor #"
    'fname = "Function Description"

    
     
     tabname = "AG3"
    On Error Resume Next
     Set sht = Worksheets(tabname)
     On Error GoTo 0
     
    
      If sht Is Nothing Then
         Worksheets.Add after:=Worksheets(Worksheets.Count)
         ActiveSheet.Name = tabname
      End If
      
     Set cell = Worksheets(tab7).Range("a1:u100").Find(fname, lookat:=xlPart)
     celladd = cell.Offset(1, 0).Address
     'Worksheets(tab7).Range(celladd, Range(celladd).End(xlDown)).Select
     'Selection.Copy
     Do Until IsEmpty(Range(celladd).Offset(i, 0))
           Worksheets(tab7).Range(celladd).Offset(i, 0).Copy
           
        If Len(Range(celladd).Offset(i, 0).Value) = 3 Then
           Worksheets(tabname).Range("a1").Offset(i, 0).PasteSpecial xlPasteAll
           With Worksheets(tabname).Range("a1").Offset(i, 0)
             .Interior.Color = RGB(32, 114, 214)
             .Font.Bold = True
             .Font.Color = RGB(255, 255, 255)
          End With
        ElseIf Len(Range(celladd).Offset(i, 0).Value) = 4 Then
            Worksheets(tabname).Range("a1").Offset(i, 0).PasteSpecial xlPasteValues
       End If
       
       Range(celladd).Offset(i, 2).Copy
       If Len(Range(celladd).Offset(i, 0).Value) = 3 Then
           Worksheets(tabname).Range("b1").Offset(i, 0).PasteSpecial xlPasteAll
           With Worksheets(tabname).Range("b1").Offset(i, 0)
             .Interior.Color = RGB(32, 114, 214)
             .Font.Bold = True
             .Font.Color = RGB(255, 255, 255)
          End With
            
        ElseIf Len(Range(celladd).Offset(i, 0).Value) = 4 Then
            Worksheets(tabname).Range("b1").Offset(i, 0).PasteSpecial xlPasteValues
            'Worksheets(tabname).Range("b1").Offset(i, 0).WrapText = True
       End If
       
       i = i + 1
     Loop
     
     Worksheets(tabname).Activate
     'Worksheets(tabname).Columns("b").WrapText = True

 
End Sub
'2 Copy function description and segment it
Sub CopyFunction()
    Dim tab7 As String
    Dim cell As Range
    Dim fname As String
    Dim i As Integer
    Dim tabname As String
    Dim sht As Worksheet
    Dim celladd As String
    Dim copycell As Range
    Dim fungroup() As String
    Dim fgroup As String
    Dim dimen As Integer
    
    
    
    'fname = "Factor #"
     fname = "Function Description"

    tab7 = "7 FactorsFunctionsPriorities"
     tabname = "AG3"
    On Error Resume Next
     Set sht = Worksheets(tabname)
     On Error GoTo 0
     
    Worksheets(tab7).Activate
    ' Add a new sheet if not existent
      If sht Is Nothing Then
         Worksheets.Add after:=Worksheets(Worksheets.Count)
         ActiveSheet.Name = tabname
      End If
      
     Set cell = Worksheets(tab7).Range("a1:u100").Find(fname, lookat:=xlPart)
     celladd = cell.Offset(1, 0).Address
     
     Do Until IsEmpty(Range(celladd).Offset(i, 0).Value)
           fgroup = Mid(Range(celladd).Offset(i, 0).Value, 2, Len(Range(celladd).Offset(i, 0).Value) - 2)
           Worksheets(tabname).Range("h1").Offset(i, 0).Value = fgroup
           
            fungroup() = Split(fgroup, ",")
            If UBound(fungroup) >= 0 And Not testarray("IL", fungroup) Then
                Debug.Print "ex/sm" & i
                Worksheets(tabname).Range("e1").Offset(i, 0) = "EX/SM"
            End If
            If testarray("IL", fungroup) Then
              Worksheets(tabname).Range("f1").Offset(i, 0) = "IL"
           End If
            
            i = i + 1
            For dimen = LBound(fungroup) To UBound(fungroup)
               Worksheets(tabname).Range("j1").Offset(i, dimen).Value = fungroup(dimen)
            Next dimen
             
     Loop
    
    Worksheets(tabname).Activate
    With Range("d1:f1")
              .Interior.Color = RGB(32, 114, 214)
             .Font.Bold = True
             .Font.Color = RGB(255, 255, 255)
    End With
    
    Range("D1").Value = "ALL"
    Range("E1").Value = "EX/SM/CM"
    Range("F1").Value = "IL"
End Sub

'---------------------3 Extract account code from file list name
Sub extractnumber()
    Dim s1 As Integer
    Dim s2 As Integer
    Dim s3 As Integer
    Dim acc_code As String
    Dim cell As Range
    Dim cnt As Integer
    
    cnt = 0
    For Each cell In Range("account")
    

    s1 = InStr(Range("a2").Offset(cnt, 0).Value, "_")
    s2 = InStr(s1 + 1, Range("a2").Offset(cnt, 0).Value, "_")
    s3 = InStr(s2 + 1, Range("a2").Offset(cnt, 0).Value, "_")
    acc_code = Mid(Range("a2").Offset(cnt, 0).Value, s2 + 1, s3 - s2 - 1)
    Range("i2").Offset(cnt, 0).Value = acc_code
    cnt = cnt + 1
    Next cell
    
End Sub
