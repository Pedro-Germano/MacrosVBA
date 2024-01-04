Sub SendEmailWithTableInClipboard()

    Dim rng As Range
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim wordDoc As Word.Document
    Dim wordRange As Word.Range
    
    ' Declare an array of strings
    Dim tableNames() As String
    Dim arraySize As Integer
    arraySize = 3
    ReDim tableNames(1 To arraySize)
    
    ' Populate the array with strings
    tableNames(1) = "LB"
    tableNames(2) = "MM"
    tableNames(3) = "RV"
    
    ' Clear references
    Set wordRange = Nothing
    Set wordDoc = Nothing
    Set olMail = Nothing
    Set olApp = Nothing
    
    ' Initialize Outlook
    Set olApp = New Outlook.Application
    ' Create a new email
    Set olMail = olApp.CreateItem(olMailItem)
    
    ' Email settings
    With olMail
        .To = "gestao@bsideinvestimentos.com"
        .Subject = "Profitability - Funds"
        
        ' Create a Word object to manipulate the email body
        Set wordDoc = olMail.GetInspector.WordEditor
        Set wordRange = wordDoc.Content
        
        ' Initialize the Word Range as demanded
        wordRange.Text = ""
        wordRange.Font.Size = wordRange.Font.Size + 5
        wordRange.Bold = True
        
        ' Loop through each element of the array
        Dim i As Integer
        For i = LBound(tableNames) To UBound(tableNames)
        
            ' Set the range of cells you want to email
            Sheets(tableNames(i) & " Valores").Select
            Range("D2").Select
            Selection.End(xlDown).Select
            Selection.End(xlDown).Select
            Selection.End(xlDown).Select
            Selection.End(xlDown).Select
            Set rng = Range(Selection, "R2")
            
            ' Copy the range to the clipboard
            rng.Copy
            
            ' Check if there is an image in the clipboard, "2" is the image format
            If Application.ClipboardFormats(2) Then
                
                ' Inserting the Title
                wordRange.Text = "___________" & vbCrLf & vbCrLf & "Fundos - " & tableNames(i) & _
                    vbCrLf & "___________" & vbCrLf
    
                ' Move the cursor to the end of the previously inserted text
                wordRange.Collapse Direction:=wdCollapseEnd
                    
                ' Paste the clipboard image into the email body
                wordRange.Paste
                
                ' Move the cursor again to the end of the inserted content
                wordRange.Collapse Direction:=wdCollapseEnd
                
            Else
                MsgBox "There is no image from '" & tableNames(i) & " Valores' in the clipboard."
                
            End If
        
        Next i
        
        ' Display the email
        .Display
    End With
    
    ' Inserting the default email signature
    Dim signatureRange As Word.Range
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    ' Retrieve the default signature
    Set signatureRange = olMail.GetInspector().WordEditor.Range
    
    ' Add the default signature to the email
    wordRange.Collapse Direction:=wdCollapseEnd
    wordRange.FormattedText = signatureRange.FormattedText
    
    ' Clean up Outlook objects
    Set olMail = Nothing
    Set olApp = Nothing
    
    ' Clear references
    Set wordRange = Nothing
    Set wordDoc = Nothing
    Set olMail = Nothing
    Set olApp = Nothing
    
End Sub



