' MOVE ALL FUNCTIONS AND SUBS INTO THIS MODULE
' ONLY HAVE EVENTS IN FORM CODE
' REMEMBER TO RENAME WITH MODULE PREFIX (MPSAdminCode.StaffSearch .... )

' Global constants and apiShowWindow - these are for the IE browser window hack
Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" _
            (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2

' Global constants - to edit based on user attachments and links - MOVE to settings?
Global Const STD_ATTACH = 0
Global Const STD_LINK = 2
Global Const DEF_BROWSER = 2
'Public SelectedBrowser As Integer

' Public constants to hold the ItemLoad() counts
Public Const strHyper As String = "LINK """
Public LoadAttMentions As Integer
Public LoadLinkMentions As Integer
Public LoadAttCount As Integer
Public LoadLinkCount As Integer
Public m_objBody As String
Public strBody As String

' Public strings to hold the URLs being fed through to OpenBrowser()
Public URLpre As String
Public URLres As String
Public WebURL As String

Public NewLinkVal As String
Public NewLinkIndex As Integer
Public NewLinkFriendly As String

' SelButton tells you WHICH button is being edited
' 1 = Top Left, 2 = Top Right, .... 6 = Bottom Right
Public SelButton As Integer

' This is an array containing B1 - 6
' Where myLinks(INDEX) = B & Index
Public myLinks(5, 5) As String
' myLinks(0) = direct.sussex.ac.uk
' myLinks(1) = sussex.ac.uk/webcontentmanager

' INITIAL VALUES/DESCRIPTION
' SelectedBrowser = 1 (IE), 2 (FF - *), 3 (Chr) ....
' B1 = BIS login ' B2 = WCM ' B3 = SxD
' B4 = SyD ' B5 = CMS ' B6 = Web Reports

' These are the User Settings variables to be read/written to the Settings.txt file
Public SelectedBrowser As Integer, B1 As String, B2 As String, B3 As String, B4 As String, B5 As String, B6 As String, StdAtt As Integer, StdLink As Integer
Public B1F As String, B2F As String, B3F As String, B4F As String, B5F As String, B6F As String

Global Const START_POS = 3

Dim IEApp As Object

' THIS IS THE POST-LOAD PRE-EDIT COUNTER FOR LINKS & ATTACHMENTS IN MESSAGES
Function MailLoadCounter(strBody As String)
' This will count the number of links etc. in the message BEFORE you start typing (i.e. replies).

' Key Declarations needed for RegExp
Set reg = CreateObject("vbscript.regexp")
On Error GoTo handleError

Dim HyperLoadCount As Integer

' END SETUP AND DECLARATIONS
' ---------------------------
' START REGEX MATCH - ATTACH MENTIONS
    With reg
        .IgnoreCase = True
        .MultiLine = True
        .pattern = "attach"
        .Global = True
    End With
Set regExp_matches = reg.Execute(strBody)
LoadAttMentions = regExp_matches.Count
' END REGEX MATCH - ATTACH MENTIONS
' -------------------------
' START REGEX MATCH - LINK MENTIONS
   With reg
        .IgnoreCase = True
        .MultiLine = True
        .pattern = "link"
        .Global = True
    End With
    Set regExp_matches = reg.Execute(strBody)
    LoadLinkMentions = regExp_matches.Count
    ' HYPERLINK CHECK
    With reg
        .IgnoreCase = False
        .Global = True
        .MultiLine = True
        .pattern = strHyper
    End With
    Set regExp_matches = reg.Execute(strBody)
    HyperLoadCount = regExp_matches.Count
    ' Works after Public Const strHyper as String
    ' MsgBox ("MLC - HyperLoadCount = " & HyperLoadCount)
    LoadLinkMentions = LoadLinkMentions - HyperLoadCount
' END REGEX MATCH - LINK MENTIONS
' ------------------------
' START REGEXP MATCH - LINK COUNTER
            With reg
                .IgnoreCase = True
                .MultiLine = True
                ' This is a full "clickable" URL checker - WORKS - although matches hyperlinks twice
                .pattern = "\b(?:(?:(?:https?|ftp|file)://|www\.|ftp\.)[-A-Z0-9+&@#/%?=~_|$!:,.;]*[-A-Z0-9+&@#/%=~_|$]|((?:mailto:)?[A-Z0-9._%+-]+@[A-Z0-9._%-]+\.[A-Z]{2,4})\b)|" & Chr(34) & "(?:(?:https?|ftp|file)://|www\.|ftp\.)[^" & Chr(34) & "\r\n]+" & Chr(34) & "?|'(?:(?:https?|ftp|file)://|www\.|ftp\.)[^'\r\n]+'?"
                .Global = True
            End With
        
            Set regExp_matches = reg.Execute(strBody)
            ' The number of links detected BEFORE editing
            LoadLinkCount = regExp_matches.Count

If LoadAttMentions = 0 And LoadLinkMentions = 0 Then
    Exit Function
Else
    'MsgBox ("Matches for LoadAttMentions = " & LoadAttMentions & vbCrLf & vbCrLf & "Matches for LoadLinkMentions = " & LoadLinkMentions)
    End If
    
handleError:
    Exit Function
    Close

End Function

Public Sub BrowserOpen(WebURL As String, BrowserID As Integer)

On Error GoTo handleError

' Quick check in case no browser selected (default = 2)
If BrowserID = 0 Then
    BrowserID = DEF_BROWSER
    'MsgBox ("Reset. BrowseID = " & BrowserID)
Else
    'MsgBox ("No Reset. BrowseID = " & BrowserID)
End If
    'MsgBox ("FirstCheck. BrowseID = " & BrowserID)

' This is where the Browser check will be carried out
            Select Case BrowserID
            
            ' N.B. IE currently DOES NOT open in a new tab. Forces a new window each time
            Case 1
                Dim pathIE As String
                pathIE = "C:\Program Files\Internet Explorer\iexplore.exe"
                    If Dir(pathIE) = "" Then
                        pathIE = "C:\Program Files\Internet Explorer\iexplore.exe"
                        End If
                    ' Is this necessary?
                    If Dir(pathIE) = "" Then
                        MsgBox "IE Path Not Found", vbCritical, "Macro Ending"
                        Exit Sub
                    End If
                    ' Is this necessary?
                    'Dim r As Long
                    Shell """" & pathIE & """" & WebURL, vbHide
                         
            Case 2
                ' --- Firefox ---
                Dim pathFireFox As String
                pathFireFox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                    If Dir(pathFireFox) = "" Then
                        pathFireFox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                        End If
                    If Dir(pathFireFox) = "" Then
                        MsgBox "FireFox Path Not Found", vbCritical, "Macro Ending"
                        Exit Sub
                    End If
                Shell """" & pathFireFox & """" & " -new-tab " & WebURL, vbHide
                ' --- Firefox ---
            
            Case 3
                ' --- Chrome ---
                Dim pathChrome As String
                pathChrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                    If Dir(pathChrome) = "" Then
                        pathChrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                        End If
                    If Dir(pathChrome) = "" Then
                        MsgBox "Chrome Path Not Found", vbCritical, "Macro Ending"
                        Exit Sub
                    End If
                Shell """" & pathChrome & """" & " -new-tab " & WebURL, vbHide
                ' --- Chrome ---
                
            Case Else
                MsgBox ("ELSE. bID = " & BrowserID)
            End Select
            
handleError:
    'MsgBox ("Error: " & vbError)
    Exit Sub
    Close
            
End Sub

Public Sub StudentSearch()

Dim StdStr As String
Dim SxDpre, SxDsuf As String

StdStr = MPSAdminWidget.StudentSearchBox.Text

' Check for blank search box
If StdStr = "" Then
    Notify.NotifyMsg.Caption = "Search Box Empty!"
    Notify.Show
    'MsgBox ("Search Box Empty!")
    Exit Sub
Else
    ' Check for numeric - i.e. Reg No to search
    If IsNumeric(StdStr) Then
                ' Leave this message as searching for incorrect reg no leads to DB errors
                MsgBox ("Searching by Reg Number - if error screen shows, check the Reg Number!")
                SxDpre = "http://sussex.ac.uk/its/sxdredirect/?page=student&rgno="
                SxDsuf = "&rel=ADMIN"
                URLres = SxDpre & StdStr & SxDsuf
                WebURL = URLres
                BrowserOpen "" & WebURL & "", SelectedBrowser
                ' Should skip the rest of this now.
            Exit Sub
            
    Else
                StdStr = LCase(StdStr)
                SxDpre = "https://direct.sussex.ac.uk/page.php?realm=searches&page=directory_search_results&formlet=student_directory&trail=directories&step=search&re_search_page=directories&surname="
                SxDsuf = "&first_name=&initials=&username=&registration_number=&dept=&current=Current&x=0&y=0"
                URLres = SxDpre & StdStr & SxDsuf
                WebURL = URLres
                BrowserOpen "" & WebURL & "", SelectedBrowser
            Exit Sub
            
    End If
    
End If

End Sub

Public Sub StaffSearch()

Dim StaffStr As String

StaffStr = MPSAdminWidget.StaffSearchBox.Text
URLpre = "http://www.sussex.ac.uk/profiles/search/"

        If StaffStr = "" Then
                Notify.NotifyMsg.Caption = "Search Box Empty!"
                Notify.Show
                'MsgBox ("Search Box Empty!")
        Else
                URLres = URLpre & StaffStr
                WebURL = URLres
                BrowserOpen "" & WebURL & "", SelectedBrowser
                Exit Sub
        End If

End Sub

Public Sub SussexSearch()
    Dim SearchTerm As String
    SearchTerm = MPSAdminWidget.SussexSearchBox.Text
    'SearchTerm = LCase(SearchTerm)
            
    SearchTerm = LCase(Replace(SearchTerm, " ", "%20"))
        
    URLpre = "http://www.sussex.ac.uk/search/?type=site&t="
    URLres = "&realm=internal"
    
        If SearchTerm = "" Then
                Notify.NotifyMsg.Caption = "Search Box Empty!"
                Notify.Show
        Else
                URLres = URLpre & SearchTerm & URLres
                WebURL = URLres
                BrowserOpen "" & WebURL & "", SelectedBrowser
                Exit Sub
        End If

End Sub

' These are the customisable button events
Public Sub WB1()
    WebURL = myLinks(1, 0)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WB2()
    WebURL = myLinks(1, 1)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WB3()
    WebURL = myLinks(1, 2)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WB4()
    WebURL = myLinks(1, 3)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WB5()
    WebURL = myLinks(1, 4)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WB6()
    WebURL = myLinks(1, 5)
    BrowserOpen "" & WebURL & "", SelectedBrowser
End Sub

Public Sub WidgetShow()
    Settings.Captions
    Load MPSAdminWidget
    MPSAdminWidget.StartUpPosition = START_POS
    MPSAdminWidget.Show vbModeless
End Sub
