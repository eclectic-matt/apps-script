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
        Case (7)
            NewLinkVal = "http://sussexfoodhospitality.com/updated/menu.aspx"
            NewLinkFriendly = "Sussex Food"
        Case (8)
            NewLinkVal = "http://www.sussex.ac.uk/mps/internal/"
            NewLinkFriendly = "MPS Internal"
        Case (9)
            NewLinkVal = "http://www.sussex.ac.uk/mps/"
            NewLinkFriendly = "MPS External"
        Case (10)
            NewLinkVal = "http://www.sussex.ac.uk/"
            NewLinkFriendly = "Sussex Homepage"
        Case (11)
            NewLinkVal = "https://webmail.sussex.ac.uk/"
            NewLinkFriendly = "Sussex Webmail"
        Case (12)
            NewLinkVal = "https://bis1.sussex.ac.uk/cognos"
            NewLinkFriendly = "Web Reports (v7)"
        Case (12)
            NewLinkVal = "https://bis1.sussex.ac.uk/cognos10"
            NewLinkFriendly = "Cognos Reports (v10)"
    End Select

myLinks(1, SelButton - 1) = NewLinkVal
myLinks(0, SelButton - 1) = NewLinkFriendly

If B1 <> "" Then B1 = myLinks(1, 0)
If B2 <> "" Then B2 = myLinks(1, 1)
If B3 <> "" Then B3 = myLinks(1, 2)
If B4 <> "" Then B4 = myLinks(1, 3)
If B5 <> "" Then B5 = myLinks(1, 4)
If B6 <> "" Then B6 = myLinks(1, 5)
If B1F <> "" Then B1F = myLinks(0, 0)
If B2F <> "" Then B2F = myLinks(0, 1)
If B3F <> "" Then B3F = myLinks(0, 2)
If B4F <> "" Then B4F = myLinks(0, 3)
If B5F <> "" Then B5F = myLinks(0, 4)
If B6F <> "" Then B6F = myLinks(0, 5)

Call Settings.Captions

Settings.SaveSettings
    
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = START_POS
End Sub

Private Sub UserForm_Initialize()
    LinkForm.StartUpPosition = START_POS
    NewLinkBox.AddItem "Sussex Direct" '0 https://direct.sussex.ac.uk/login.php?realm=home/
    NewLinkBox.AddItem "Study Direct" '1 https://studydirect.sussex.ac.uk/login/
    NewLinkBox.AddItem "Business Applications" '2 http://www.sussex.ac.uk/its/services/staffservices/businessapplications/
    NewLinkBox.AddItem "CMS - Yellow Screens" '3 http://owf.admin.sussex.ac.uk/jbisapps/
    NewLinkBox.AddItem "Web Content Manager" '4 http://www.sussex.ac.uk/wcm/entry/
    NewLinkBox.AddItem "Agresso" '5 https://abw.admin.sussex.ac.uk/Agresso/System/Login.aspx
    NewLinkBox.AddItem "Google Calendar" '6 https://www.google.com/calendar/
    NewLinkBox.AddItem "Sussex Food" '7 http://sussexfoodhospitality.com/updated/menu.aspx
    NewLinkBox.AddItem "MPS Internal" '8 http://www.sussex.ac.uk/mps/internal/
    NewLinkBox.AddItem "MPS External" '9 http://www.sussex.ac.uk/mps/
    NewLinkBox.AddItem "Sussex homepage" '10 http://www.sussex.ac.uk/
    NewLinkBox.AddItem "Sussex Webmail" '11 https://webmail.sussex.ac.uk/
    NewLinkBox.AddItem "Web Reports (v7)" '12 https://bis1.sussex.ac.uk/cognos
    NewLinkBox.AddItem "Cognos Reports (v10)" '13 https://bis1.sussex.ac.uk/cognos10
End Sub

Private Sub UserForm_Terminate()
    MPSAdminWidget.Show vbModeless
End Sub
