Option Compare Database
Option Explicit

'There are two text box built in Access form that are named "textfilepath","textfilename" respectively

'1 First Link to the excel file we want to import

Private Sub LinkData_Click()
          
          Dim item As Variant
          Dim filestr As String
          Dim fname As String
          Dim tblname As String
          With Application.FileDialog(msoFileDialogFilePicker)
              .AllowMultiSelect = False
              .Filters.Add "Excel files", "*.xlsx"
              
              If .Show Then
                For Each item In .SelectedItems
                     Me.TextFilepath = item
                     filestr = Me.TextFilepath
                     
                     fname = Right(filestr, Len(filestr) - InStrRev(filestr, "\"))

                     On Error Resume Next
                     tblname = Left(fname, InStr(fname, " ") - 1)
                     Debug.Print fname
                     On Error GoTo 0
                     If Err.Number <> 0 Then
                        tblname = Left(fname, InStr(fname, "_") - 1)
                     End If
                      Me.TextFileName = tblname
                     
                Next item
             End If
         End With
               
          
End Sub

Private Sub ImportData_Click()
         Dim objxl As New Excel.Application
         Dim wbk As Excel.Workbook
         Dim wsht As Object
         
         Set wbk = objxl.Workbooks.Open(Me.TextFilepath)
         
         For Each wsht In wbk.Worksheets
              If Not wsht.Name Like "*Pivot" Then
              
                DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8 _
                , Me.TextFileName & wsht.Name, Me.TextFilepath, True, wsht.Name & "$"
              End If
        Next wsht
        
    
    wbk.Close
    Set wbk = Nothing
    
    objxl.Quit
    
    Set objxl = Nothing
         
End Sub
