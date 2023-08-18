Const MACRO_NAME = "OST2XLS"
Const MSG_CODE_LENGTH = 7

Dim excApp As Object, _
    excWkb As Object, _
    excWks As Object, _
    intVersion As Integer, _
    intMessages As Integer, _
    lngRow As Long, _
    lngSmtps As Long
    
'Diccionario smtp adresses
Dim dict_smtp_addr As Object 'Declare a generic Object reference




Sub ExportMessagesToExcel()
    Dim strFilename As String, olkSto As Outlook.Store
    Dim count_stores As Integer
    Set dict_smtp_addr = CreateObject("Scripting.Dictionary") 'Late Binding of the Dictionary
    'strFilename = InputBox("Enter a filename (including path) to save the exported messages to.", MACRO_NAME)
    strFilename = "C:\Users\luism\OneDrive\_FruitSpec\LMCD_messages6.xlsx"
    If strFilename <> "" Then
        intMessages = 0
        intVersion = GetOutlookVersion()
        Set excApp = CreateObject("Excel.Application")
        excApp.ReferenceStyle = xlA1
        Set excWkb = excApp.Workbooks.Add
        count_stores = 1
        For Each olkSto In Session.Stores
            Set excWks = excWkb.Worksheets.Add()
            excWks.Name = "LMCD_messages_" & Str(count_stores)
            count_stores = count_stores + 1
            'Write Excel Column Headers
            With excWks
                .cells(1, 1) = "msgCode"
                .cells(1, 2) = "msgFolder"
                .cells(1, 3) = "msgSender"
                .cells(1, 4) = "msgReceived"
                .cells(1, 5) = "msgSent To"
                .cells(1, 6) = "msgCC"
                .cells(1, 7) = "msgBCC"
                .cells(1, 8) = "msgSubject"
                .cells(1, 9) = "msgRecipients"
            End With
            lngRow = 2
            Call ProcessFolder(olkSto.GetRootFolder(), "root")
        Next
        'add sheet for smtp addresses
        Set excWks = excWkb.Worksheets.Add()
        excWks.Name = "SMTP_Addresses"
        With excWks
                .cells(1, 1) = "smtpAddress"
                .cells(1, 2) = "smtpDomain"
        End With
        Call Process_dict_smtp_addr
        excWkb.SaveAs strFilename
        
    End If
    Set excWks = Nothing
    Set excWkb = Nothing
    excApp.Quit
    Set excApp = Nothing
    MsgBox "Process complete.  A total of " & intMessages & " messages were exported. And " & lngSmtps & " smtps.", vbInformation + vbOKOnly, "Export messages to Excel"
End Sub

Sub ProcessFolder(olkFld As Outlook.MAPIFolder, str_folder As String)
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    Dim strMessages As String
    Dim arrRecipients() As String
    Dim strSender As String
    'Write messages to spreadsheet
    For Each olkMsg In olkFld.Items
        'Only export messages, not receipts or appointment requests, etc.
        If olkMsg.Class = olMail Then
            'Add a row for each field in the message you want to export
            'yield message code
            arrRecipients = Split(GetSMTPAddressForRecipients(olkMsg), ";")
            strMessages = Trim(Str(intMessages))
            strMessages = "'" & String(MSG_CODE_LENGTH - Len(strMessages), "0") & strMessages
            strSender = GetSMTPAddress(olkMsg, intVersion)
            
            'insert sender in dictionary
            insert_dict_smtp_addr (strSender)
            
            
            For Each reci In arrRecipients
                excWks.cells(lngRow, 1) = strMessages
                excWks.cells(lngRow, 2) = str_folder & "\" & olkFld.Name
                excWks.cells(lngRow, 3) = strSender
                excWks.cells(lngRow, 4) = olkMsg.ReceivedTime
                excWks.cells(lngRow, 5) = olkMsg.To
                excWks.cells(lngRow, 6) = olkMsg.CC
                excWks.cells(lngRow, 7) = olkMsg.BCC
                excWks.cells(lngRow, 8) = olkMsg.Subject
                excWks.cells(lngRow, 9) = reci
                lngRow = lngRow + 1
            Next
            intMessages = intMessages + 1
        End If
    Next
    Set olkMsg = Nothing
    For Each olkSub In olkFld.Folders
        Call ProcessFolder(olkSub, str_folder & "\" & olkFld.Name)
    Next
    Set olkSub = Nothing
End Sub

Private Function GetSMTPAddress(Item As Outlook.MailItem, intOutlookVersion As Integer) As String
    Dim olkSnd As Outlook.AddressEntry, olkEnt As Object
    On Error Resume Next
    Select Case intOutlookVersion
        Case Is < 14
            If Item.SenderEmailType = "EX" Then
                GetSMTPAddress = SMTP2007(Item)
            Else
                GetSMTPAddress = Item.SenderEmailAddress
            End If
        Case Else
            Set olkSnd = Item.Sender
            If olkSnd.AddressEntryUserType = olExchangeUserAddressEntry Then
                Set olkEnt = olkSnd.GetExchangeUser
                GetSMTPAddress = olkEnt.PrimarySmtpAddress
            Else
                GetSMTPAddress = Item.SenderEmailAddress
            End If
    End Select
    On Error GoTo 0
    Set olkPrp = Nothing
    Set olkSnd = Nothing
    Set olkEnt = Nothing
End Function
Function GetSMTPAddressForRecipients(mail As Outlook.MailItem) As String
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim contact As Outlook.ContactItem
    Dim resultado As String
    Dim lngLenResultado As Long
    Dim subresultado As String
    Dim smtp_mail As String
    
    'On Error Resume Next

    Set recips = mail.Recipients
    reusltado = ""
    For Each recip In recips
        If recip.Resolved Then
            Select Case recip.AddressEntry.AddressEntryUserType
                Case Is = olExchangeAgentAddressEntry              '3  An address entry that is an Exchange agent.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olExchangeDistributionListAddressEntry  '1   An address entry that is an Exchange distribution list.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olExchangeOrganizationAddressEntry      '4   An address entry that is an Exchange organization.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olExchangePublicFolderAddressEntry      '2   An address entry that is an Exchange public folder.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olExchangeRemoteUserAddressEntry        '5   An Exchange user that belongs to a different Exchange forest.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olExchangeUserAddressEntry              '0   An Exchange user that belongs to the same Exchange forest.
                    smtp_mail = recip.AddressEntry.GetExchangeUser.PrimarySmtpAddress
                    If smtp_mail = "" Then
                        subresultado = recip.Name & "(0-WITHOUT SMTP)"
                    Else
                        subresultado = smtp_mail
                    End If
                    
                Case Is = olLdapAddressEntry                      '20  An address entry that uses the Lightweight Directory Access Protocol (LDAP).
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olOtherAddressEntry                     '40  A custom or some other type of address entry such as FAX.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olOutlookContactAddressEntry            '10  An address entry in an Outlook Contacts folder.
                    'Set contact = recip.AddressEntry.GetContact()
                    'If contact Is Nothing Then
                        subresultado = recip.Address
                    'Else
                    '    subresultado = recip.AddressEntry.GetContact.Email1Address & ";"
                    'End If
                Case Is = olOutlookDistributionListAddressEntry   '11  An address entry that is an Outlook distribution list.
                    subresultado = recip.Name & "(" & Str(recip.AddressEntry.AddressEntryUserType) & ")"
                Case Is = olSmtpAddressEntry                      '30  An address entry that uses the Simple Mail Transfer Protocol (SMTP).
                    subresultado = recip.Address & ""
                Case Else
                    subresultado = recip.Name & "(else)"
            End Select
        Else
            subresultado = recip.Name & "(else)"
        End If
        
        'Insert recipient smtp in the dictionary
        insert_dict_smtp_addr subresultado
        
        resultado = resultado & subresultado & ";"
        
    Next
    
        GetSMTPAddressForRecipients = removelastchar(resultado)
End Function


Function GetOutlookVersion() As Integer
    Dim arrVer As Variant
    arrVer = Split(Outlook.Version, ".")
    GetOutlookVersion = arrVer(0)
End Function

Function SMTP2007(olkMsg As Outlook.MailItem) As String
    Dim olkPA As Outlook.PropertyAccessor
    On Error Resume Next
    Set olkPA = olkMsg.PropertyAccessor
    SMTP2007 = olkPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")
    On Error GoTo 0
    Set olkPA = Nothing
End Function
Sub Process_dict_smtp_addr()
Dim smtp_addr As Variant
Dim domain As String
Dim lngRowSmtp As Long
    lngRowSmtp = 2 ' first line is for headers
    For Each smtp_addr In dict_smtp_addr
        domain = dict_smtp_addr(smtp_addr)
        excWks.cells(lngRowSmtp, 1) = smtp_addr
        excWks.cells(lngRowSmtp, 2) = domain
        lngRowSmtp = lngRowSmtp + 1
    Next smtp_addr
End Sub
Function domain(smtp_addr As String) As String
Dim intPosArroba As Integer
Dim intLenSmtp_addr As Integer
Dim resultado As String
    intLenSmtp_addr = Len(smtp_addr)
    intPosArroba = InStr(smtp_addr, "@")
    If intPosArroba = 0 Then
        'not found. as output we return input
        domain = smtp_addr
    Else
        domain = Right(smtp_addr, intLenSmtp_addr - intPosArroba)
    End If
End Function
Function removelastchar(texto As String) As String
Dim lnglentexto As Integer

    'remove trailing semicolon
    If texto <> "" Then
        lnglentexto = Len(texto)
        removelastchar = Left(texto, lnglentexto - 1)
    Else
        removelastchar = texto
    End If

End Function

Sub insert_dict_smtp_addr(key As String)
    If Not dict_smtp_addr.exists(key) Then
        dict_smtp_addr.Add key, domain(key)
        lngSmtps = lngSmtps + 1
    End If
End Sub

