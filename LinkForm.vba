Private Sub Cancel_Click()
    LinkForm.Hide
    MPSAdminWidget.Show vbModeless
End Sub

Private Sub Confirm_Click()
    Call NewLink
    LinkForm.Hide
    MPSAdminWidget.Show vbModeless
End Sub

Private Sub NewLinkBox_AfterUpdate()
    Call NewLink
    LinkForm.Hide
    MPSAdminWidget.Show vbModeless
End Sub

Private Sub NewLink()

'Link = Value
'Index# = Index
NewLinkIndex = NewLinkBox.ListIndex
    
    Select Case (NewLinkIndex)
        Case (0)
            NewLinkVal = "https://direct.sussex.ac.uk/login.php?realm=home/"
            NewLinkFriendly = "Sussex Direct"
        Case (1)
            NewLinkVal = "https://studydirect.sussex.ac.uk/login/"
            NewLinkFriendly = "Study Direct"
        Case (2)
            NewLinkVal = "http://www.sussex.ac.uk/its/services/staffservices/businessapplications/"
            NewLinkFriendly = "BIS Login Page"
        Case (3)
            NewLinkVal = "http://owf.admin.sussex.ac.uk/jbisapps/"
            NewLinkFriendly = "CMS"
        Case (4)
            NewLinkVal = "http://www.sussex.ac.uk/wcm/entry/"
            NewLinkFriendly = "Web Content Manager"
        Case (5)
            NewLinkVal = "https://abw.admin.sussex.ac.uk/Agresso/System/Login.aspx"
            NewLinkFriendly = "Agresso"
        Case (6)
            NewLinkVal = "https://www.google.com/calendar/"
            NewLinkFriendly = "Google Calendar"
    End Select
    Notify.NotifyMsg.Caption = "New Link set as " & NewLinkFriendly
    Notify.Show

myLinks(1, SelButton) = NewLinkVal
If B1 <> "" Then B1 = myLinks(1, 0)
If B2 <> "" Then B2 = myLinks(1, 1)
If B3 <> "" Then B3 = myLinks(1, 2)
If B4 <> "" Then B4 = myLinks(1, 3)
If B5 <> "" Then B5 = myLinks(1, 4)
If B6 <> "" Then B6 = myLinks(1, 5)

'Notify.NotifyMsg.Caption = "LINK FORM" & vbCrLf & _
        "SelBrow = " & SelectedBrowser & vbCrLf & _
        "B1 = " & B1 & vbCrLf & _
        "B2 = " & B2 & vbCrLf & _
        "B3 = " & B3 & vbCrLf & _
        "B4 = " & B4 & vbCrLf & _
        "B5 = " & B5 & vbCrLf & _
        "B6 = " & B6 & vbCrLf & _
        "StdAtt = " & StdAtt & vbCrLf & _
        "StdLink = " & StdLink
'Notify.Show

'MsgBox ("LINK FORM" & vbCrLf & _
'        "SelBrow = " & SelectedBrowser & vbCrLf & _
'        "B1 = " & B1 & vbCrLf & _
'        "B2 = " & B2 & vbCrLf & _
'        "B3 = " & B3 & vbCrLf & _
'        "B4 = " & B4 & vbCrLf & _
'        "B5 = " & B5 & vbCrLf & _
'        "B6 = " & B6 & vbCrLf & _
'        "StdAtt = " & StdAtt & vbCrLf & _
'        "StdLink = " & StdLink)

Settings.SaveSettings
    
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = START_POS
End Sub

Private Sub UserForm_Initialize()
    LinkForm.StartUpPosition = START_POS
    NewLinkBox.AddItem "Sussex Direct" '0 'https://direct.sussex.ac.uk/login.php?realm=home/
    NewLinkBox.AddItem "Study Direct" '1 'https://studydirect.sussex.ac.uk/login/
    NewLinkBox.AddItem "Business Applications" '2 'http://www.sussex.ac.uk/its/services/staffservices/businessapplications/
    NewLinkBox.AddItem "CMS - Yellow Screens" '3 'http://owf.admin.sussex.ac.uk/jbisapps/
    NewLinkBox.AddItem "Web Content Manager" '4 'http://www.sussex.ac.uk/wcm/entry/
    NewLinkBox.AddItem "Agresso" '5 'https://abw.admin.sussex.ac.uk/Agresso/System/Login.aspx
    NewLinkBox.AddItem "Google Calendar" '6 'https://www.google.com/calendar/
End Sub

Private Sub UserForm_Terminate()
    MPSAdminWidget.Show vbModeless
End Sub


