VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDuplicates 
   Caption         =   "Duplicates"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Selection"
      Height          =   360
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "&Merge"
      Height          =   360
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7235
      Width           =   1305
   End
   Begin MSDBGrid.DBGrid dbgDuplicates 
      Bindings        =   "Duplicates.frx":0000
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "Duplicates.frx":001C
      TabIndex        =   3
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Data datDuplicates 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboSearchType 
      Height          =   315
      ItemData        =   "Duplicates.frx":1405
      Left            =   6360
      List            =   "Duplicates.frx":1407
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1305
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2412
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   10
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label lblFoundNumber 
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the account you wish to keep, then click merge"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6720
      Width           =   10575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search Type:"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Criteria :-"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmDuplicates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngMasterAccountNum As Long
Dim mlngToBeMergeAccountNum As Long
Dim mstrMasterAccountName As String
Dim mstrToBeMergeAccountName As String
Dim lstrScreenHelpFile As String

Public Sub cmdBack_Click()

    Me.Enabled = False
    gstrButtonRoute = gconstrMainMenu
    UnloadLastForm
    Set gstrCurrentLoadedForm = frmMainReps
    mdiMain.DrawButtonSet gstrButtonRoute
    Me.Enabled = True
    frmMainReps.Show
    
End Sub

Private Sub cmdClear_Click()

    mlngMasterAccountNum = 0
    mlngToBeMergeAccountNum = 0
    mstrMasterAccountName = ""
    mstrToBeMergeAccountName = ""

End Sub

Private Sub cmdFind_Click()
Dim lstrSQL As String
Dim lintRetVal As Integer

    If Trim$(txtSearchCriteria) = "" Then
        MsgBox "You enter criteria to search upon!", vbInformation, gconstrTitlPrefix & "Duplicate Handling"
        Exit Sub
    Else
        Select Case cboSearchType
        Case "TelephoneNum"
            If Len(Trim(txtSearchCriteria)) < 4 Then
                lintRetVal = MsgBox("WARNING: Searching on such a narrow criteria may take some time!" & _
                    vbCrLf & vbCrLf & "Do you wish to proceed?", vbYesNo, gconstrTitlPrefix & "Duplicate Handling")
                If lintRetVal <> vbYes Then
                    Exit Sub
                End If
            End If
        End Select
    End If
    lstrSQL = "SELECT CustNum, Trim(Trim([Salutation]) & ' ' & Trim([Initials])" & _
        "& ' ' & Trim([Surname])) AS Name, Trim([Add1] & ', ') & Trim([Add2] & ', ')" & _
        "& Trim([Add3] & ', ') & Trim([Add4] & ', ') & Trim([Add5] & ', ')" & _
        "& Trim([PostCode]) AS Address, TelephoneNum, EveTelephoneNum, " & _
        "EMail, AccountType, AcctStatus FROM CustAccounts where " & _
        cboSearchType & " like '*" & Trim(txtSearchCriteria) & "*' ;"
        
    datDuplicates.RecordSource = lstrSQL
    datDuplicates.Refresh
    
    lblFoundNumber = "Found " & datDuplicates.Recordset.RecordCount & " records."

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdMerge_Click()
Dim llngChosenAccountNum As Long
Dim lintRetval2 As Integer
Dim lintRetVal As Integer
Dim lstrSQL As String

    With gstrCustomerAccount
        On Error Resume Next
        llngChosenAccountNum = CLng(datDuplicates.Recordset("CustNum"))
           
        If Err.Number = 3021 Then
            MsgBox "You must select a Customer!", , gconstrTitlPrefix & "Duplicate Handling"
            Exit Sub
        End If
        On Error GoTo ErrHandler
        
        If mlngMasterAccountNum = 0 Then
            mstrMasterAccountName = datDuplicates.Recordset("Name")
        
            lintRetval2 = MsgBox("You have selected a Master account :-" & vbCrLf & _
                vbTab & "(" & llngChosenAccountNum & ") '" & mstrMasterAccountName & "'" & _
                vbCrLf & vbCrLf & _
                "Is this correct ?" & vbCrLf & vbCrLf & _
                "(If so, Select an Account you wish to merge to it!)", _
                vbYesNo + vbInformation, gconstrTitlPrefix & "Duplicate Handling")
            If lintRetval2 = vbYes Then
                mlngMasterAccountNum = llngChosenAccountNum
            End If
            Exit Sub
        ElseIf mlngToBeMergeAccountNum = 0 Then
            mlngToBeMergeAccountNum = llngChosenAccountNum
            mstrToBeMergeAccountName = datDuplicates.Recordset("Name")
            
            If mlngToBeMergeAccountNum = mlngMasterAccountNum Then
                MsgBox "You may not merge an Account to itself!" & vbCrLf & _
                "This selection has been clear, but your Master account has not!" & vbCrLf & vbCrLf & _
                "Please select another!", vbInformation, gconstrTitlPrefix & "Duplicate Handling"
                
                mlngToBeMergeAccountNum = 0
                mstrToBeMergeAccountName = ""
                Exit Sub
            End If
            lintRetVal = MsgBox("Do you wish to merge all orders " & vbCrLf & _
                "From :- " & vbCrLf & vbTab & "(" & mlngToBeMergeAccountNum & ") '" & _
                mstrToBeMergeAccountName & "'" & vbCrLf & vbCrLf & _
                "To   :-   (Master)" & vbCrLf & vbTab & _
                "(" & mlngMasterAccountNum & ") '" & mstrMasterAccountName & "'" & vbCrLf & _
                vbCrLf & "Are you sure you wish to proceed?" & vbCrLf & _
                "(This process will also delete customer record (" & mlngToBeMergeAccountNum & "))", vbQuestion + vbYesNo, _
                gconstrTitlPrefix & "Duplicate Handling")
            If lintRetVal = vbYes Then
                Busy True, Me
                ShowStatus 110: DoEvents
                lstrSQL = "UPDATE AdviceNotes SET CustNum = " & mlngMasterAccountNum & _
                    " WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL
                
                ShowStatus 111: DoEvents
                lstrSQL = "UPDATE CustNotes SET CustNum = " & mlngMasterAccountNum & _
                    " WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL

                ShowStatus 112: DoEvents
                lstrSQL = "UPDATE Cashbook SET CustNum = " & mlngMasterAccountNum & _
                    " WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL

                ShowStatus 113: DoEvents
                lstrSQL = "UPDATE OrderLinesMaster SET CustNum = " & mlngMasterAccountNum & _
                    " WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL

                ShowStatus 114: DoEvents
                lstrSQL = "UPDATE Pforce SET CustNum = " & mlngMasterAccountNum & _
                    " WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL

                ShowStatus 115: DoEvents
                
                Sleep 2000
                ShowStatus 116: DoEvents
                lstrSQL = "DELETE * From CustAccounts WHERE (((CustNum)=" & mlngToBeMergeAccountNum & "));"
                gdatCentralDatabase.Execute lstrSQL

                ShowStatus 117: DoEvents
                Busy False, Me
                MsgBox "Process Complete!", vbInformation, gconstrTitlPrefix & "Duplicate Handling"
            Else
                MsgBox "All selections have been cleared!", , gconstrTitlPrefix & "Duplicate Handling)"
            End If
        End If

        mlngMasterAccountNum = 0
        mlngToBeMergeAccountNum = 0
        mstrMasterAccountName = ""
        mstrToBeMergeAccountName = ""
        
    End With

Exit Sub
ErrHandler:
    
    Busy False, Me
    
    Select Case GlobalErrorHandler(Err.Number, "frmDuplicates.cmdMerge_Click", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
        
    ShowBanner Me
    
    cboSearchType.AddItem "Postcode"
    cboSearchType.AddItem "TelephoneNum"
    cboSearchType.ListIndex = 0
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datDuplicates.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datDuplicates.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
    datDuplicates.RecordSource = "select * from CustAccounts where 1=0"
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()

    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With


    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With

    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdMerge
        .Top = Me.Height - gconlongButtonTop
        .Left = (cmdBack.Left - .Width) - 120
    End With
    
    With cmdClear
        .Top = Me.Height - gconlongButtonTop
        .Left = (cmdMerge.Left - .Width) - 120
    End With
        
    With lblFoundNumber
        .Top = (cmdHelp.Top - .Height) - 305
    End With
    
    With lblInstructions
        .Left = 0
        .Width = Me.Width
        .Top = lblFoundNumber.Top + 120
    End With
    
    With dbgDuplicates
        .Width = Me.Width - 360
        If (cmdHelp.Top - .Top) > 665 Then
            .Height = (cmdHelp.Top - .Top) - 665
        Else
            .Height = 665 - (cmdHelp.Top - .Top)
        End If
    End With
    
End Sub

Private Sub txtSearchCriteria_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Duplicates.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_DEDUPE_MAIN
    ctlBanner1.WhatIsID = IDH_DEDUPE_MAIN

    txtSearchCriteria.WhatsThisHelpID = IDH_DEDUPE_SEACRIT
    cmdFind.WhatsThisHelpID = IDH_STANDARD_FIND
    cboSearchType.WhatsThisHelpID = IDH_DEDUPE_SEATYPE
    dbgDuplicates.WhatsThisHelpID = IDH_DEDUPE_GRIDDUPS
    cmdMerge.WhatsThisHelpID = IDH_DEDUPE_MERGE
    cmdClear.WhatsThisHelpID = IDH_DEDUPE_CLEAR
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
