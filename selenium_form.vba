Option Explicit

Public IsReady As Boolean
Public Status As String
Public IsReactivated As Boolean
Public IsBlocked As Boolean
Public IsExcluded As Boolean
Public IsDeleted As Boolean
Public IsActive As Boolean
Public IsFreeTrial As Boolean
Public IsEdgeCase As Boolean
Public SearchAccessMode As Boolean


Private Sub cmdLoadURLs_Click()
    Dim thisSheet As Worksheet
    IsReady = True
    ActiveSheet.Cells(1, 7) = "BOOYA!"
End Sub


Private Sub GrantSearch_Click()
    SearchAccessMode = True
End Sub

Private Sub Reactivate_Click()
    IsReactivated = True
End Sub

Private Sub Reset_Click()
    ActiveSheet.Cells(1, 5) = 1
    ActiveSheet.Cells(1, 6) = "closed"
End Sub

Private Sub UserBlocked_Click()
    IsBlocked = True
End Sub

Private Sub Exclude_Click()
    IsExcluded = True
End Sub

Private Sub StartProcess_Click()
    Dim LatestCounter As Integer, thisSheet As Worksheet
    Dim TemplateURL As String, ThisSelenium As New SeleniumWrapper.WebDriver
    Dim ThisURL As String, FF As Object, keys As New SeleniumWrapper.keys
    Dim LoginURL As String, VerifyBlockText As String, UserURL As String
    Dim IsRegionBlocked As Boolean
    
    'reset all controls
    IsReactivated = False
    IsExcluded = False
    IsBlocked = False
    SearchAccessMode = False
    'set sheet and URL elements
    Set thisSheet = ActiveSheet
    TemplateURL = "https://www.connectifier.com"
    UserURL = "/admin/user/"
    'sync counter and status
    LatestCounter = thisSheet.Cells(1, 5)
    Status = thisSheet.Cells(1, 6)
    
    'main process is a while loop which terminates when status is set to finished
    Do While Status <> "finished"
        'this is why this macro has a chance of running
        DoEvents
        If IsReady Then
            ThisURL = thisSheet.Cells(LatestCounter, 1)
            'this must run first if browser has not been initialized
            If Status = "closed" And ThisURL <> "" Then
                'initialize browser and log in
                Set ThisSelenium = StartURL(ThisSelenium, TemplateURL)
                Set ThisSelenium = OpenURL(ThisSelenium, "/enter-email")
                Set ThisSelenium = TypeInLogin(ThisSelenium)
                Set ThisSelenium = TypeInPassword(ThisSelenium)
                'status set
                thisSheet.Cells(1, 6) = "open"
                LatestCounter = LatestCounter + 1
            'do this if browser is initialized and log in is successful
            ElseIf ThisURL <> "" And Status = "open" Then
                'directs page to Admin user page
                Set ThisSelenium = OpenURL(ThisSelenium, UserURL & ThisURL)
                'get user status and toggle controls
                Select Case GetUserStatus(ThisSelenium)
                    Case "TRIAL_EXPIRED"
                        IsReactivated = True
                    Case "BLOCKED"
                        IsBlocked = True
                    Case "DELETED"
                        IsDeleted = True
                    Case "ACCOUNT_ACTIVE"
                        IsActive = True
                    Case "FREE_TRIAL"
                        IsFreeTrial = True
                    Case "ACCOUNT_EXCLUDED"
                        IsExcluded = True
                    Case Else
                        IsEdgeCase = True
                End Select
                
            ElseIf ThisURL = "" Then
                thisSheet.Cells(1, 6) = "finished"
            End If
                'reset by deactivating main path
                IsReady = False
                thisSheet.Cells(1, 7) = "nah"
        End If
        
        'if we are granting search access to the user
        If SearchAccessMode Then
            ThisURL = thisSheet.Cells(LatestCounter, 1)
            If ThisURL <> "" Then
                Set ThisSelenium = OpenURL(ThisSelenium, UserURL & ThisURL)
                Call GrantSearchAccess(ThisSelenium)
                LatestCounter = LatestCounter + 1
                thisSheet.Cells(1, 7) = "nah"
            End If
        End If
        
        'mangle of conditionals to direct responses to user status
        If IsReactivated Then
            If CheckIfBlocked(ThisSelenium) Then
                thisSheet.Cells(LatestCounter, 2) = GetBlockStatus(ThisSelenium)
            Else
                Call ClickButton(ThisSelenium, "FREE_TRIAL")
                thisSheet.Cells(LatestCounter, 2) = "Reactivated"
            End If
            IsReactivated = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsExcluded Then
            thisSheet.Cells(LatestCounter, 2) = "Excluded"
            IsExcluded = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsBlocked Then
            If CheckIfBlocked(ThisSelenium) Then
                thisSheet.Cells(LatestCounter, 2) = GetBlockStatus(ThisSelenium)
            Else
                thisSheet.Cells(LatestCounter, 2) = "Blocked"
            End If
            IsBlocked = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsActive Then
            thisSheet.Cells(LatestCounter, 2) = "Active"
            IsActive = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsDeleted Then
            thisSheet.Cells(LatestCounter, 2) = "Deleted"
            IsDeleted = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsFreeTrial Then
            thisSheet.Cells(LatestCounter, 2) = "Reactivated"
            IsFreeTrial = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        If IsEdgeCase Then
            thisSheet.Cells(LatestCounter, 2) = "ERROR"
            IsEdgeCase = False
            LatestCounter = LatestCounter + 1
            IsReady = True
        End If
        'sync counter and status
        Status = thisSheet.Cells(1, 6)
        thisSheet.Cells(1, 5) = LatestCounter
    Loop
    'when everything's done!
    ThisSelenium.Close
    MsgBox ("finished!")
End Sub