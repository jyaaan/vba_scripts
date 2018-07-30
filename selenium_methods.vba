
Function StartURL( _
     ThisSelenium As SeleniumWrapper.WebDriver, _
     URL As String, _
     Optional Timeout As Long = 120000, _
     Optional ImplicitWait As Long = 5000) As SeleniumWrapper.WebDriver
    Call ThisSelenium.Start("chrome", URL)
    ThisSelenium.setTimeout (Timeout)
    ThisSelenium.setImplicitWait (ImplicitWait)
    ThisSelenium.waitForPageToLoad 12000
    Set StartURL = ThisSelenium
End Function

Function OpenURL( _
     ThisSelenium As SeleniumWrapper.WebDriver, _
     URL As String, _
     Optional Timeout As Long = 120000, _
     Optional ImplicitWait As Long = 5000) As SeleniumWrapper.WebDriver
    ThisSelenium.Open URL, Timeout
    ThisSelenium.waitForPageToLoad 12000
    Set OpenURL = ThisSelenium
End Function


Sub ClickButton( _
     ThisSelenium As SeleniumWrapper.WebDriver, _
     ButtonToClick As String, _
     Optional PauseDur As Long = 10000)
    Dim ThisButton As WebElement
    ThisSelenium.waitForPageToLoad 12000
    Set ThisButton = ThisSelenium.findElementByXPath("//button[@value='FREE_TRIAL']")
    ThisButton.Click
    ThisSelenium.waitForPageToLoad 12000
End Sub

Sub GrantSearchAccess( _
     ThisSelenium As SeleniumWrapper.WebDriver)

    Dim ThisButton As WebElement
    ThisSelenium.waitForPageToLoad 12000
    Set ThisButton = ThisSelenium.findElementByXPath("//button[@value='SEARCH_ACCESS']")
    If ThisButton.getAttribute("class") <> "btn btn-sm  btn-primary " Then
        ThisButton.Click
        ThisSelenium.waitForPageToLoad 12000
    End If

End Sub

Function SearchForText( _
     ThisSelenium As SeleniumWrapper.WebDriver, _
     TextToSearch As String) As Boolean

    Dim ClassString As String
    ThisSelenium.waitForPageToLoad 12000
    ClassString = ThisSelenium.findElementByXPath("//button[@value='BLOCKED']").getAttribute("class")
    If ClassString = "btn btn-sm  btn-primary " Then
        SearchForText = True
    Else
        SearchForText = False
    End If
    
End Function

Function CheckIfRegionBlocked( _
     ThisSelenium As SeleniumWrapper.WebDriver) As Boolean
    Dim result As String
    ThisSelenium.waitForPageToLoad 12000
    result = ThisSelenium.verifyTextPresent("BLOCKED_REGION")
    If result = "OK" Then
        CheckIfRegionBlocked = True
    Else
        CheckIfRegionBlocked = False
    End If
End Function

Function CheckIfIPBlocked( _
     ThisSelenium As SeleniumWrapper.WebDriver) As Boolean
    Dim result As String
    ThisSelenium.waitForPageToLoad 12000
    result = ThisSelenium.verifyTextPresent("BLOCKED_IP")
    If result = "OK" Then
        CheckIfIPBlocked = True
    Else
        CheckIfIPBlocked = False
    End If
End Function

Function GetBlockStatus( _
     ThisSelenium As SeleniumWrapper.WebDriver) As String

    If CheckIfRegionBlocked(ThisSelenium) Then
        GetBlockStatus = "RegionBlock"
    ElseIf CheckIfIPBlocked(ThisSelenium) Then
        GetBlockStatus = "IPBlock"
    End If

End Function

Function CheckIfBlocked( _
     ThisSelenium As SeleniumWrapper.WebDriver) As Boolean
    CheckIfBlocked = CheckIfRegionBlocked(ThisSelenium) Or CheckIfIPBlocked(ThisSelenium)
End Function
Function GetUserStatus( _
     ThisSelenium As SeleniumWrapper.WebDriver) As String
    Dim ClassString As String
    ThisSelenium.waitForPageToLoad 12000
    ClassString = ThisSelenium.findElementByXPath("//button[@class='btn btn-sm  btn-primary ']").getAttribute("value")
    GetUserStatus = ClassString
End Function

Function TypeInLogin( _
     ThisSelenium As SeleniumWrapper.WebDriver) As SeleniumWrapper.WebDriver
    Dim SubmitButton As WebElement
    Set SubmitButton = ThisSelenium.findElementByXPath("//button[@type='submit']")
    ThisSelenium.Type "name=email", "jyamashiro@connectifier.com"
    SubmitButton.Click
    ThisSelenium.waitForPageToLoad 12000
    Set TypeInLogin = ThisSelenium
End Function
Function TypeInPassword( _
     ThisSelenium As SeleniumWrapper.WebDriver) As SeleniumWrapper.WebDriver
    Dim SubmitButton As WebElement
    ThisSelenium.waitForPageToLoad 12000
    Set SubmitButton = ThisSelenium.findElementByXPath("//button[@type='submit']")
    ThisSelenium.Type "name=password", (Chr(52) & Chr(56) & Chr(111) & Chr(114) & Chr(105) & Chr(104) & Chr(115) & Chr(97) & Chr(109) & Chr(97) & Chr(89))
    SubmitButton.Click
    ThisSelenium.waitForPageToLoad 12000
    Set TypeInPassword = ThisSelenium
End Function