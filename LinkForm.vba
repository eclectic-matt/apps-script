Private Sub UserForm_Deactivate()
    NewLinkBox.Clear
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
    NewLinkBox.AddItem "MPS NEW DEV" '14 wwwnewdev.sussex.ac.uk/mps/internal/
    NewLinkBox.AddItem "Student Lists" '15 https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=student_programme_lists
    NewLinkBox.AddItem "Module Search" '16 https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=course_directory
    NewLinkBox.AddItem "Course Search" '17 https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=programme_directory
    NewLinkBox.AddItem "ITS - Report a fault" '18 http://www.sussex.ac.uk/its/help/smartform
    NewLinkBox.AddItem "ITS - Staff Help" '19 http://www.sussex.ac.uk/its/staff
    NewLinkBox.AddItem "Room Bookings" '20 http://www.sussex.ac.uk/roombooking/view_timetable.php
    NewLinkBox.AddItem "Broadcast" '21 http://www.sussex.ac.uk/broadcast/dashboard
    NewLinkBox.AddItem "B&H Buses" '22 http://www.buses.co.uk/travel/live-bus-times.aspx
    NewLinkBox.AddItem "National Rail - Falmer Station" '23 http://ojp.nationalrail.co.uk/service/ldbboard/dep/FMR
    NewLinkBox.AddItem "SPA Office - Exam Handbooks" '24 http://www.sussex.ac.uk/academicoffice/documentsandpolicies/examinationandassessmenthandbooks
    NewLinkBox.AddItem "MPS Dev Site Dashboard" '25 http://www.sussex.ac.uk/wcm/dashboard/page/structure/441
    ' REMEMBER TO ADD TO CASE-SELECT STATEMENT IN THE NewLink() SUB
    
End Sub

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
            NewLinkVal = "https://bis1.sussex.ac.uk/cognos/"
            NewLinkFriendly = "Web Reports (v7)"
        Case (13)
            NewLinkVal = "https://bis1.sussex.ac.uk/cognos10/"
            NewLinkFriendly = "Cognos Reports (v10)"
        Case (14)
            NewLinkVal = "wwwnewdev.sussex.ac.uk/mps/internal/"
            NewLinkFriendly = "MPS New Dev Site"
        Case (15)
            NewLinkVal = "https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=student_programme_lists"
            NewLinkFriendly = "Student Lists"
        Case (16)
            NewLinkVal = "https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=course_directory"
            NewLinkFriendly = "Module Directory"
        Case (17)
            NewLinkVal = "https://direct.sussex.ac.uk/page.php?realm=searches&page=directories&formlet=programme_directory"
            NewLinkFriendly = "Course Directory"
        Case (18)
            NewLinkVal = "http://www.sussex.ac.uk/its/help/smartform/"
            NewLinkFriendly = "ITS - Report a fault"
        Case (19)
            NewLinkVal = "http://www.sussex.ac.uk/its/staff/"
            NewLinkFriendly = "ITS - Staff Help"
        Case (20)
            NewLinkVal = "http://www.sussex.ac.uk/roombooking/view_timetable.php"
            NewLinkFriendly = "GTS Room Bookings"
        Case (21)
            NewLinkVal = "http://www.sussex.ac.uk/broadcast/dashboard"
            NewLinkFriendly = "Broadcast"
        Case (22)
            NewLinkVal = "http://www.buses.co.uk/travel/live-bus-times.aspx"
            NewLinkFriendly = "B&H Buses"
        Case (23)
            NewLinkVal = "http://ojp.nationalrail.co.uk/service/ldbboard/dep/FMR/"
            NewLinkFriendly = "National Rail - Falmer"
        Case (24)
            NewLinkVal = "http://www.sussex.ac.uk/academicoffice/documentsandpolicies/examinationandassessmenthandbooks/"
            NewLinkFriendly = "SPA - Exam Handbook"
        Case (25)
            NewLinkVal = "http://www.sussex.ac.uk/wcm/dashboard/page/structure/441/"
            NewLinkFriendly = "MPS New Dev Site - Dashboard"
        Case Default
            NewLinkVal = "http://www.sussex.ac.uk/"
            NewLinkFriendly = "Link Not Found - Sussex Homepage"
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

Private Sub UserForm_Terminate()
    MPSAdminWidget.Show vbModeless
End Sub
