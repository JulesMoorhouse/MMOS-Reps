Attribute VB_Name = "modReports"
Option Explicit

Sub Main()
Dim lintDebugVersion As Variant
Dim lstrThisHelpFile As String

    gdatSystemStartTime = Now()
    
    gstrSystemRoute = srStandardRoute
    
    lstrThisHelpFile = MCLDebugChoices
        
    SetSystemNames
    
    frmSplash.Show
    frmSplash.Refresh
    
    If InStr(UCase(Command$), "/X") = 0 Then
        MsgBox "You cannot run the program from here!" & vbCrLf & _
            "You must use the Loader program!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    End If
    
    Select Case CheckForOtherMMosprog(frmSplash.hwnd)
    Case True
        MsgBox "You may only run one " & gconstrProductFullName & " program at once!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    Case False
        'MsgBox "no other prog found!"
    End Select
    
    InitDb
    
    frmLogin.Show vbModal
    If Not frmLogin.OK Then
        'Login Failed so exit app
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
    
        UpdateLoader
        Unhook
        End
    End If

    If gstrGenSysInfo.lngUserLevel < 30 Then 'Less than Sales
        MsgBox "You do not have security rights to run MMOS Manager!" & vbCrLf & vbCrLf & _
            "Please contact your IT Support Office!", vbInformation, gconstrTitlPrefix & "Startup"
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
    
        UpdateLoader
        Unhook
        End
    End If
    Busy True
   
    ConcurrencyTest
    
    UpdateLists frmSplash, True 
    
    UserLicenceCheck
    
    gbooJustPreLoading = True
    
    ShowStatus 0
    DoEvents

    Unload frmLogin
    DoEvents
            
    Load frmChildCalendar
    Unload frmChildCalendar
    Load frmChildGenericDropdown
    Unload frmChildGenericDropdown
    Load frmChildLabelOptions
    Unload frmChildLabelOptions
    Load frmChildOptions
    Unload frmChildOptions
    Load frmChildPrinter
    Unload frmChildPrinter
    Load frmChildStaMultiAdd
    Unload frmChildStaMultiAdd
    Load frmDuplicates
    Unload frmDuplicates
    Load frmPForce
    Unload frmPForce
    Load frmPrintPreview
    Unload frmPrintPreview
    Load frmReportOptions
    Unload frmReportOptions
    Load frmReports
    Unload frmReports
    Load frmStaMultiMarket
    Unload frmStaMultiMarket
    Load frmSummary
    Unload frmSummary
    
    gbooJustPreLoading = False
    
    CopyHelpFile lstrThisHelpFile
    
    Load frmMainReps

    Unload frmSplash

    gstrVATRate = gstrReferenceInfo.strVATRate175 ' Does change with overseas flag

    frmMainReps.Show
    
    Busy False
    
End Sub
Sub UpdateReportingData(pobjForm As Form)
Dim lstrSourcePath As String
Dim lstrDestinationPath As String
Dim lstrSourceFile As String
Dim lstrDestinationFile As String


    Busy True, pobjForm
    
    ShowStatus 38
    DoEvents
    
    lstrSourcePath = gstrStatic.strServerPath
    lstrDestinationPath = gstrStatic.strServerPath & "Output\"
    
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrSourceFile = gstrStatic.strCentralTestingDBFile
    Else
        lstrSourceFile = gstrStatic.strCentralDBFile
    End If
    
    
    lstrDestinationFile = lstrDestinationPath & "YestDay.mdb"
    
    On Error GoTo ErrHandler
    If DateValue(Date) > DateValue(FileDateTime(lstrDestinationFile)) Then
        
        CopyFile lstrSourceFile, lstrDestinationFile
        ShowStatus 99
        DoEvents
    End If
    Busy False, pobjForm

Exit Sub
ErrHandler:

    Select Case Err.Number
    Case 70  'permission denied, someone in the database
        Busy False, pobjForm
        MsgBox "THIS MESSAGE ONLY APPLIES TO EXTERNAL ACCESS REPORTS, NOT INTERNAL REPORTS (Which use live data). " & vbCrLf & _
        "Someone is using the reporting database, preventing it from being updated!" & vbCrLf & _
            "This is not a technical problem with TMOS!" & vbCrLf & vbCrLf & _
            "Please contact you Technical support office!" & vbCrLf & _
            "WARNING: Until this problem is rectified you will be working with " & vbCrLf & _
            "out of date data!", , gconstrTitlPrefix & "Reporting Data"
        Exit Sub
    Case 53
        Resume Next
    End Select
    Select Case GlobalErrorHandler(Err.Number, "UpdateReportingData", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        Exit Sub
    Case Else
        Resume Next
    End Select
    
End Sub

Sub CopyCentralToRepsDB(pstrFilename As String, pstrCopyString As String)
Dim lintFileNum As Integer

    lintFileNum = FreeFile

    Open pstrFilename For Append As lintFileNum
    
    Print #lintFileNum, "@echo Please wait, until prompted!"
    Print #lintFileNum, "@" & pstrCopyString
    Print #lintFileNum, "@echo You may now close this window!"
    
    Close #lintFileNum

End Sub
