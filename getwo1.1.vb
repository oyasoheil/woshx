Sub ImportWOColumnsWithDebug()
    Dim folderPath As String
    Dim fileName As String
    Dim currentFile As String
    Dim wbWO As Workbook
    Dim wsWO As Worksheet
    Dim wsDest As Worksheet
    Dim wsDebug As Worksheet
    Dim destRow As Long
    Dim debugRow As Long
    Dim jVal As Variant, lVal As Variant, nVal As Variant
    Dim rowOffset As Long
    Dim fileCount As Long
    Dim dataCopied As Boolean
    Dim headerExtracted As Boolean

    ' === Create or get destination sheet (Data) ===
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Data")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = "Data"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0

    ' === Create or clear Debug sheet ===
    On Error Resume Next
    Set wsDebug = ThisWorkbook.Sheets("Debug")
    If wsDebug Is Nothing Then
        Set wsDebug = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDebug.Name = "Debug"
    Else
        wsDebug.Cells.Clear
    End If
    On Error GoTo 0

    ' Set starting rows
    destRow = 2
    debugRow = 2
    headerExtracted = False

    ' Add debug headers
    wsDebug.Range("A1:E1").Value = Array("File Name", "Status", "Rows Copied", "Error", "Timestamp")

    ' === Folder selection ===
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing WO Files"
        If .Show <> -1 Then
            MsgBox "No folder selected!", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1) & "\"
    End With

    Application.ScreenUpdating = False
    fileCount = 0
    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""
        currentFile = Left(fileName, Len(fileName) - 5)
        dataCopied = False

        On Error GoTo ErrorHandler
        Set wbWO = Workbooks.Open(folderPath & currentFile, ReadOnly:=True)
        Set wsWO = wbWO.Sheets(1)

        ' === Only once: Extract J7, L7, N7 into B1, C1, D1 ===
        If Not headerExtracted Then
            wsDest.Range("B1").Value = wsWO.Range("J7").Value
            wsDest.Range("C1").Value = wsWO.Range("L7").Value
            wsDest.Range("D1").Value = wsWO.Range("N7").Value
            headerExtracted = True
        End If

        rowOffset = 1 ' Start from J8, L8, N8
        Dim rowsCopiedForFile As Long
        rowsCopiedForFile = 0

        Do
            jVal = wsWO.Range("J7").Offset(rowOffset, 0).Value
            lVal = wsWO.Range("L7").Offset(rowOffset, 0).Value
            nVal = wsWO.Range("N7").Offset(rowOffset, 0).Value

            If IsEmpty(jVal) And IsEmpty(lVal) And IsEmpty(nVal) Then Exit Do
            wsDest.Cells(destRow, "A").Value = currentFile
            wsDest.Cells(destRow, "B").Value = jVal
            wsDest.Cells(destRow, "C").Value = lVal
            wsDest.Cells(destRow, "D").Value = nVal
            

            destRow = destRow + 1
            rowOffset = rowOffset + 1
            rowsCopiedForFile = rowsCopiedForFile + 1
            dataCopied = True
        Loop

        wbWO.Close SaveChanges:=False

        ' Log success
        wsDebug.Cells(debugRow, "A").Value = currentFile
        wsDebug.Cells(debugRow, "B").Value = "Success"
        wsDebug.Cells(debugRow, "C").Value = rowsCopiedForFile
        wsDebug.Cells(debugRow, "E").Value = Now
        debugRow = debugRow + 1
        fileCount = fileCount + 1
        GoTo ContinueLoop

ErrorHandler:
        wsDebug.Cells(debugRow, "A").Value = currentFile
        wsDebug.Cells(debugRow, "B").Value = "Error"
        wsDebug.Cells(debugRow, "C").Value = 0
        wsDebug.Cells(debugRow, "D").Value = Err.Description
        wsDebug.Cells(debugRow, "E").Value = Now
        debugRow = debugRow + 1
        Err.Clear
        Resume ContinueLoop

ContinueLoop:
        fileName = Dir
    Loop

    Application.ScreenUpdating = True
    MsgBox fileCount & " file(s) processed. Check 'Data' and 'Debug' sheets.", vbInformation
End Sub


