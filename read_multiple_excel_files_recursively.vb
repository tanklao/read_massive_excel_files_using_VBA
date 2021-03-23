Sub recursiveGet()
'
'This script is developed by Tank Liu at Northern Illinois University (tanklao((at))gmail.com)
'
    Dim filePath As String
    filePath = "C:\Users\" & Environ$("username") & "\Documents\Dev\SampleData\WpYToJdpDi679\Newfolder001"
    Dim colFiles As New Collection
    
    Open filePath & ".log" For Output As #500 'open a text file to log process
    
    'get all xls file names to a collection called colFiles we just defined
    RecursiveDir colFiles, filePath, "*.xls", True
    
    'loop with the files we got
    Dim vfile As Variant
    num = 1
    For Each vfile In colFiles
        Debug.Print "Processing file No. "; num; "..."
        Call processWorkbook(CStr(vfile))
        num = num + 1
    Next vfile
    
    'add hyperlink to file URI
    For x = 2 To ActiveSheet.UsedRange.Rows.Count + 1
        If InStr(1, Cells(x, 1).Value, ":", vbTextCompare) > 1 Then
            ActiveSheet.Hyperlinks.Add Anchor:=Cells(x, 1), Address:=Cells(x, 1).Value, TextToDisplay:=Cells(x, 1).Value
        End If
        If x >= 66530 Then Exit For
    Next x
    ActiveWorkbook.Save
    Close #500
    Debug.Print "All done! "; num; "files processed!"
End Sub

Sub processWorkbook(fName As String)
    Dim resultSheet As Worksheet
    Set resultSheet = ActiveSheet
    Dim rowBegin, colNum As Integer
     'find the first empty cell on column A
    rowBegin = resultSheet.UsedRange.Rows.Count + 1
    Dim openBook As Workbook
    Set openBook = Workbooks.Open(fName)
    Dim wksht As Worksheet
    Dim coord As Variant
        
    'Proccess each work sheet
    'This part is where you need to change according to your own task
    For i = 1 To openBook.Sheets.Count
        Debug.Print rowBegin; " ";
        Set wksht = openBook.Sheets(i)
        
        Debug.Print fName; vbTab; wksht.Name;
        Print #500, fName; vbTab; wksht.Name;
        coord = findDataRangeFromWorksheet(wksht, "Last Name", "Bio")
        
        resultSheet.Cells(rowBegin, 1).Value = fName
        resultSheet.Cells(rowBegin, 2).Value = wksht.Name
        resultSheet.Cells(rowBegin, 3).Value = wksht.Range( _
            wksht.Cells(coord(1), coord(2) + 1), _
            wksht.Cells(coord(3), coord(4) + 1)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        colNum = 4
        For j = coord(1) To coord(3)
            resultSheet.Cells(rowBegin, colNum).Value = wksht.Cells(j, coord(2) + 1).Value
            colNum = colNum + 1
        Next j
        rowBegin = rowBegin + 1
        Debug.Print vbTab; "done"
        Print #500, vbTab; "done"
    Next i
    openBook.Close SaveChanges:=False
    resultSheet.Activate
    Application.ScreenUpdating = True
End Sub

'This function finds the data range you need.
'wkst is the worksheet you would like to work
'The function returns an array with 4 elements corresponding to a range you can select
'first search text should appear on the left upper to the second search text
Function findDataRangeFromWorksheet(wksht As Worksheet, firstSearchText As String, secondSearchText As String) As Variant
    Dim rng As Range
    Dim r0, c0, rn, cn As Integer
    Dim arr(1 To 4) As Variant
    Set rng = wksht.Cells.Find(What:=firstSearchText, After:=Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If rng Is Nothing Then
        r0 = 1
        co = 1
    Else
        r0 = rng.Row
        c0 = rng.Column
    End If
    Set rng = wksht.Cells.Find(What:=secondSearchText, After:=Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If rng Is Nothing Then
        rn = 1
        cn = 1
    Else
        rn = rng.Row
        cn = rng.Column
    End If
    'Debug.Print r0; c0; rn; cn
    arr(1) = r0
    arr(2) = c0
    arr(3) = rn
    arr(4) = cn
    findDataRangeFromWorksheet = arr
End Function


'This function collects file names of the file type you specify from the folder.
'Don't change anything inside it. You don't have to understand the whole codes.
'You just need to know the input and output.
' ====Explanation of the parameters=====
' ByRef colFiles As Collection -- This collection is used to collect the file names.
'                                   It is an output and it's passed by reference
' ByVal strFolder As String -- here you input the foler you want to be processed
' ByVal strFileSpec As String -- here you can specify the file type you would like to find.
'                                   It does not necessarily to be a file extension
' ByVal bIncludeSubfolders As Boolean -- If you want to find files recursively then set True here

Public Function RecursiveDir(ByRef colFiles As Collection, _
                             ByVal strFolder As String, _
                             ByVal strFileSpec As String, _
                             ByVal bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function

'This function appends "\" to path that does not end with "\"
Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

