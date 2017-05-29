Attribute VB_Name = "SCM_Module"
Sub SCM()
    'Get token from input
    Dim token As Variant
    token = InputBox("Enter a token to search for in the current Workbook. (Every row containing the token will be saved to a new Workbook)")
    
    'Get the Book Name
    Dim curWbName, outWbName As String
    curWbName = Application.ActiveWorkbook.FullName
    outWbName = Replace(curWbName, ".xlsm", " " & token & ".xlsm")
    
    Dim fso As FileSystemObject
    
    'Check if the SCM Workbook file exists
    'if so, delete it
    Set fso = New FileSystemObject
    If (fso.FileExists(outWbName)) Then
        fso.DeleteFile (outWbName)
    End If
    
    'Store input Workbook
    Dim inWb As Workbook
    Set inWb = Application.ActiveWorkbook
    
    'Create a new SCM Workbook
    Dim outWb As Workbook
    Set outWb = Application.Workbooks.Add()
    'Create output worksheet with name "SCM_" + token + "_sheet"
    Dim outWsName As String
    outWsName = "SCM_" & token & "_sheet"
    outWb.Worksheets("Sheet1").Name = outWsName
        
    'Call the function to output all relevant rows in
    'the output Workbook
    Call SCM_Work(token, outWsName, inWb, outWb)
    
    outWb.SaveAs outWbName, 52
    outWb.Close
End Sub

Sub SCM_Work(token, outWsName, inWb, outWb)
    Dim curWs, outWs As Worksheet
    Set outWs = outWb.Worksheets(outWsName)
    Dim curOutRowNum, curInRowNum As Integer
    Dim curRow As Range
    curOutRowNum = 1
    
    'Loop through worksheets in current workbook to find tokens
    For Each curWs In inWb.Worksheets
        'Loop through each row in current worksheet to find tokens
        For curInRowNum = 1 To curWs.UsedRange.Rows.Count
            Set curRow = curWs.UsedRange.Rows(curInRowNum)
            
            'Stop at first empty row with no NOT DATA value
            If (curRow.Find("*") Is Nothing) Then
                Debug.Print "Finished Execution."
                Exit For
            End If
            
            'Match token in the row
            If (Not (curRow.Cells.Find(token & "*") Is Nothing)) Then
                Call copyRow(curWs, outWs, curInRowNum, curOutRowNum)
                curOutRowNum = curOutRowNum + 1
            End If
        Next
    Next
    
End Sub

' Function to copy a row from one worksheet to another
Sub copyRow(inSheet, outSheet, inRowNum, outRowNum)
    Debug.Print ("copying row")
    inSheet.Rows(inRowNum).Copy
    outSheet.Rows(outRowNum).PasteSpecial (8)
    outSheet.Rows(outRowNum).PasteSpecial (-4163)
    outSheet.Rows(outRowNum).PasteSpecial (-4122)
End Sub
