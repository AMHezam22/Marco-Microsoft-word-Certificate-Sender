Const xlUp As Long = -4162
Sub GenerateAndEmailCertificates()

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim wordDoc As Document
    Dim outlookApp As Object
    Dim outlookMail As Object

    Dim i As Long
    Dim lastRow As Long
    Dim personName As String
    Dim personEmail As String
    Dim pdfPath As String
    Dim folderPath As String

    ' Set folder for saving PDFs
    folderPath = "C:\Certificates\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    ' Open Excel workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open("The Path to your file\NamesList.xlsx")
    Set xlWs = xlWb.Sheets(1)
    lastRow = xlWs.Cells(xlWs.Rows.Count, 1).End(xlUp).Row

    ' Set Word and Outlook references
    Set wordDoc = ActiveDocument
    Set outlookApp = CreateObject("Outlook.Application")

    ' Loop over each row in Excel
    For i = 2 To lastRow

        personName = xlWs.Cells(i, 1).Value
        personEmail = xlWs.Cells(i, 2).Value

        If personName <> "" And personEmail <> "" Then

            ' Replace [NAME] with actual name inside text boxes
            Call ReplaceInTextBoxes(wordDoc, "[NAME]", personName)

            ' Export certificate to PDF
            pdfPath = folderPath & Replace(personName, " ", "_") & "_Certificate.pdf"
            wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF

            ' Create Outlook email with attachment
            Set outlookMail = outlookApp.CreateItem(0)
            With outlookMail
                .To = personEmail
                .Subject = "Your Certificate"
                .Body = "Dear " & personName & "," & vbCrLf & vbCrLf & _
                        "Please find your certificate attached." & vbCrLf & vbCrLf & "Best regards,"
                .Attachments.Add pdfPath
                .Send ' use .Display if you want to preview the email
            End With

            ' Revert the name back to [NAME]
            Call ReplaceInTextBoxes(wordDoc, personName, "[NAME]")

        End If

    Next i

    ' Cleanup
    xlWb.Close False
    xlApp.Quit

    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set outlookApp = Nothing
    Set outlookMail = Nothing

    MsgBox "All certificates generated and sent!"

End Sub




Sub ReplaceInTextBoxes(doc As Document, findText As String, replaceText As String)
    Dim shp As Shape
    Dim txtRange As Range

    For Each shp In doc.Shapes
        If shp.Type = msoTextBox Then
            Set txtRange = shp.TextFrame.TextRange

            With txtRange.Find
                .Text = findText
                .Replacement.Text = replaceText
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .Execute Replace:=wdReplaceAll
            End With

            ' Make replacement bold
            txtRange.Find.Execute FindText:=replaceText
            If txtRange.Find.Found Then
                txtRange.Font.Bold = True
            End If
        End If
    Next shp
End Sub
