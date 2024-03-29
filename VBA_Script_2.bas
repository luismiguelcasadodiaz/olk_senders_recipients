Dim excApp As Object, _
    excWkb As Object, _
    excWks As Object, _
    intMessages As Integer, _
    lngRow As Long, _
    lngSmtps As Long

Sub walk_folders()
Dim olkFld As Outlook.MAPIFolder

Set excApp = CreateObject("Excel.Application")
excApp.ReferenceStyle = xlA1
Set excWkb = excApp.Workbooks.Add
Set excWks = excWkb.Worksheets.Add()
excWks.Name = "Folders"
lngRow = 1
intMessages = 0
For Each olkSto In Session.Stores
    Call procesaCarpeta(olkSto.GetRootFolder(), "\" & olkSto.GetRootFolder().Name)
Next
excWkb.SaveAs "C:\Users\luism\OneDrive\_FruitSpec\folders4.xlsx"
excApp.Quit
MsgBox "Process complete.  A total of " & intMessages & " messages were exported. And " & lngSmtps & " smtps.", vbInformation + vbOKOnly, "Export messages to Excel"

End Sub

Sub procesaCarpeta(olkFld As Outlook.MAPIFolder, str_folder As String)
Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    'Debug.Print olkFld.Name, olkFld.Items.Count
For Each olkMsg In olkFld.Items
    'Debug.Print str_folder & "\\" & olkFld.Name, olkFld.Items.Count, olkMsg.Class
    intMessages = intMessages + 1
Next

Set olkMsg = Nothing
For Each olkSub In olkFld.Folders
        'Debug.Print str_folder & "\" & olkSub.Name, olkSub.Folders.Count, olkSub.Class
        With excWks
                .cells(lngRow, 1) = str_folder & "\" & olkSub.Name
                .cells(lngRow, 2) = olkSub.Folders.Count
                .cells(lngRow, 3) = olkSub.Class
                .cells(lngRow, 4) = olkSub.Items.Count

        End With
        lngRow = lngRow + 1
        Call procesaCarpeta(olkSub, str_folder & "\" & olkSub.Name)
Next
Set olkSub = Nothing
End Sub
