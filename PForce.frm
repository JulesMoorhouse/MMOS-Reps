VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPForce 
   Caption         =   "Please Select an Account..."
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ClipControls    =   0   'False
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
      TabIndex        =   17
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdPrintThermals 
      Caption         =   "&Print Thermals"
      Height          =   360
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdDumpSetup 
      Caption         =   "&Dump Setup"
      Height          =   360
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrintManifest 
      Caption         =   "&Print Manifest"
      Height          =   360
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdCreateFile 
      Caption         =   "&Create E.File"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Frame fraSearchBy 
      Caption         =   "Search By"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
      Begin VB.ComboBox cboConsignStatus 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Consignment Status"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   7
         Tag             =   "Status"
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Consi&gnment Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "Consign"
         Top             =   240
         Width           =   1812
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Customer &Name"
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Tag             =   "Name"
         Top             =   480
         Width           =   1812
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Order Number"
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Tag             =   "OrderNum"
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Customer Number"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "CustNumber"
         Top             =   480
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2412
   End
   Begin MSDBGrid.DBGrid dbgConsignments 
      Bindings        =   "PForce.frx":0000
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "PForce.frx":001E
      TabIndex        =   9
      Top             =   2280
      Width           =   10245
   End
   Begin VB.Data datConsignments 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PForce"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRePrint 
      Caption         =   "&RePrint Thermal"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   2160
      Top             =   1680
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   19
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Criteria :-"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmPForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrConsignStatus() As String
Dim mstrRoute As String
Dim mfrmCallingForm As Object
Dim mbooDumpedSetup As Boolean
Dim mstrLPTPortNumber As String
Dim lstrScreenHelpFile As String

Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Public Property Let CallingForm(pstrCallingForm As Object)

    Set mfrmCallingForm = pstrCallingForm

End Property
Public Property Get CallingForm() As Object

    CallingForm = mfrmCallingForm
    
End Property
Private Sub cboConsignStatus_Click()

    optSearchField(4).Value = True
    optSearchField_Click (4)
    
End Sub

Public Sub cmdBack_Click()

    If mfrmCallingForm.Name = "frmMainReps" Then
        gstrButtonRoute = gconstrMainMenu
    ElseIf mfrmCallingForm.Name = "frmReports" Then
        gstrButtonRoute = gconstrGenralReporting
    End If
      
    mdiMain.DrawButtonSet gstrButtonRoute
    Unload Me

    Set gstrCurrentLoadedForm = mfrmCallingForm
    mfrmCallingForm.Show

End Sub

Private Sub cmdCreateFile_Click()
Dim lintRetVal As Variant

    Busy True, Me
    datConsignments.Refresh

    If PForceFileGeneral(gstrStatic.strPFElecFile) = False Then
        Busy False, Me
        Exit Sub
    End If
    
    Busy False, Me
    
    lintRetVal = MsgBox("Do you wish to FLAG items as Downloaded? " & vbCrLf & _
        "If so, click YES!", vbYesNo, gconstrTitlPrefix & "PF File Creation")
    If lintRetVal = vbYes Then
        Busy True, Me
        UpdatePForceStatus "D", "P"
        datConsignments.Refresh
        Busy False, Me
    End If
    
End Sub

Private Sub cmdDumpSetup_Click()
Dim lintFileNum As Integer
Dim lstrFileName As String

    If mbooDumpedSetup = False Then
        GetPrinterInfo mstrLPTPortNumber, 66, frmChildPrinter
            
        Busy True, Me
        lintFileNum = FreeFile
        lstrFileName = GetTempDir & "T" & Format(Now(), "MMDDSSN")
        
        Open lstrFileName & ".bat" For Append As lintFileNum
        
        Print #lintFileNum, "type " & Chr(34) & gstrStatic.strServerPath & "PForce\Font.fnt" & _
            Chr(34) & " > " & gstrLPTPortNumber & ":"
        Print #lintFileNum, "type " & Chr(34) & gstrStatic.strServerPath & "PForce\Pffont.bin" & _
            Chr(34) & " > " & gstrLPTPortNumber & ":"
        Print #lintFileNum, "type " & Chr(34) & gstrStatic.strServerPath & "PForce\pfgrap.bin" & _
            Chr(34) & " > " & gstrLPTPortNumber & ":"
                
        Close #lintFileNum
        
        Busy False, Me
            
        RunNWait lstrFileName & ".bat"
        DoEvents
        On Error Resume Next
        Kill lstrFileName & ".bat"
        mbooDumpedSetup = True
    Else
        MsgBox "You have already dumped printer setup to printer!", , gconstrTitlPrefix & "Dump Setup"
    End If
    
End Sub

Private Sub cmdFind_Click()
Const lstrEndOfMessage = ", or select a different search method"
Dim lstrExtraSQL As String

    If optSearchField(0).Value = True Then  ' Consignment Number
        If Trim$(txtSearchCriteria) = "" Then
            MsgBox "You must enter a Consignment Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where ConsignNum = '" & Trim$(txtSearchCriteria) & "';"
        End If
    ElseIf optSearchField(1).Value = True Then ' Customer Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Customer Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where CustNum = " & CLng(txtSearchCriteria) & ";"
        End If
    ElseIf optSearchField(2).Value = True Then ' Order Number
        If CLng(Val(txtSearchCriteria)) = 0 Then
            MsgBox "You must enter a Order Number" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where OrderNum = " & CLng(txtSearchCriteria) & ";"
        End If
    ElseIf optSearchField(3).Value = True Then ' Customer name
        If Trim$(txtSearchCriteria) = "" Then
            MsgBox "You must enter a Customer Name" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
            lstrExtraSQL = "where DeliverySurname = '" & Trim$(txtSearchCriteria) & "';"
        End If
    ElseIf optSearchField(4).Value = True Then ' Consignment Status
        If Trim$(cboConsignStatus) = "" Then
            MsgBox "You must select a Status" & lstrEndOfMessage, , gconstrTitlPrefix & "Searching"
            Exit Sub
        Else
        
            lstrExtraSQL = "where Status = '" & Trim$(NotNull(cboConsignStatus, lstrConsignStatus)) & "';"
        End If
    End If

    Busy True, Me
    datConsignments.RecordSource = "select * from PForce " & lstrExtraSQL
    datConsignments.Refresh
    dbgConsignments.Refresh
    Busy False, Me
    
End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub
Private Sub cmdPrintManifest_Click()
Dim lstrFileName As String
Dim lintRetVal As Variant
Dim lintRetval2 As Variant
    

    lintRetval2 = MsgBox("Would you like to use the improved Manifest?", vbYesNo, gconstrTitlPrefix & "Manifest Print")
    If lintRetval2 = vbYes Then
        datConsignments.Refresh

        ChooseLayout ltParcelForceManifest, Me
        PrintObjPForceManifestGeneral
       
        ShowPlotReport
    Else
        GetPrinterInfo mstrLPTPortNumber, 66, frmChildPrinter
        Busy True, Me
        lstrFileName = GetTempDir & "m" & Format(Now(), "MMDDSSN")
    
        datConsignments.Refresh
        PForceManifestGeneral lstrFileName & ".tmp"
        Busy False, Me
        
        BatchFile lstrFileName
    End If
    
End Sub

Private Sub cmdPrintThermals_Click()
Dim lintRetVal As Variant
Dim lstrFileName As String
    
    datConsignments.Refresh
    If mbooDumpedSetup = False Then
        lintRetVal = MsgBox("Would you like to Dump Thermal Printer setting to printer?", _
            vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Print Thermals")
        If lintRetVal = vbYes Then
            
            cmdDumpSetup_Click
        Else
            GetPrinterInfo mstrLPTPortNumber, 66, frmChildPrinter
        End If
    End If

    Busy True, Me
    lstrFileName = GetTempDir & "T" & Format(Now(), "MMDDSSN")
    GetPForceAwaitings lstrFileName & ".tmp"
    
    BatchFile lstrFileName
    
    Busy False, Me
    lintRetVal = MsgBox("Did your Thermal labels Print correctly? " & vbCrLf & _
        "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Print Thermals")
    If lintRetVal = vbYes Then
        Busy True, Me
        UpdatePForceStatus "P", "A"
        datConsignments.Refresh
        Busy False, Me
    End If
  
End Sub

Private Sub cmdRePrint_Click()
Dim llngOrderNum As Long
Dim llngCustomerNum As Long
Dim lintRetVal As Variant
Dim lstrConsignmentNumber As String
Dim lstrFileName As String

    dbgConsignments.Col = 1
    
    If Trim$(dbgConsignments.Text) = "" Then
        MsgBox "no consignments shown to select from!", , gconstrTitlPrefix & "Re-Print"
        Exit Sub
    End If
    
    llngOrderNum = CLng(dbgConsignments.Text)
    
    dbgConsignments.Col = 0
    llngCustomerNum = CLng(dbgConsignments.Text)
    
    dbgConsignments.Col = 3
    lstrConsignmentNumber = dbgConsignments.Text
    
    lintRetVal = MsgBox("Do you wish to Re-Print a Thermal Label for " & vbCrLf & _
        "Consignment Number & " & lstrConsignmentNumber & " ?", vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Re-Print")
    
    datConsignments.Refresh
    
    If lintRetVal = vbYes Then
        If mbooDumpedSetup = False Then
            lintRetVal = MsgBox("Would you like to Dump Thermal Printer setting to printer?", vbYesNo + vbDefaultButton1, gconstrTitlPrefix & "Re-Print")
            If lintRetVal = vbYes Then
                
                cmdDumpSetup_Click
            Else
                GetPrinterInfo mstrLPTPortNumber, 66, frmChildPrinter
            End If
        End If
    
        Busy True, Me
        lstrFileName = GetTempDir & "T" & Format(Now(), "MMDDSSN")
        gstrPForceServiceInd.strListName = "PForce Service Indicator"
        
        GetPForceConsignment llngOrderNum, llngCustomerNum, lstrConsignmentNumber
        gstrPForceServiceInd.strListCode = gstrPForceConsignment.strServiceID
        GetListVarsAll gstrPForceServiceInd
        PFPCLFile lstrFileName & ".tmp"
        
        BatchFile lstrFileName
    End If
    
End Sub

Private Sub dbgConsignments_ButtonClick(ByVal ColIndex As Integer)

    Select Case ColIndex
    Case 2
        frmChildOptions.List = "Consignment Status"
        frmChildOptions.Code = dbgConsignments.Columns(2).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(2).Value = frmChildOptions.Code
    Case 4
        frmChildOptions.List = "PForce Service Indicator"
        frmChildOptions.Code = dbgConsignments.Columns(4).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(4).Value = frmChildOptions.Code
    Case 18
        frmChildOptions.List = "PForce Weekend Handling Code"
        frmChildOptions.Code = dbgConsignments.Columns(18).Value
        frmChildOptions.Show vbModal
        If Trim$(frmChildOptions.Code) = "" Then
            dbgConsignments.Columns(18).Value = " "
        Else
            dbgConsignments.Columns(18).Value = frmChildOptions.Code
        End If
    Case 19
        frmChildOptions.List = "PForce Prepaid Indicator"
        frmChildOptions.Code = dbgConsignments.Columns(19).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(19).Value = frmChildOptions.Code
    Case 20
        frmChildOptions.List = "PForce Notification Code"
        frmChildOptions.Code = dbgConsignments.Columns(20).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(20).Value = frmChildOptions.Code
    Case 22
        frmChildOptions.List = "Y or N"
        frmChildOptions.Code = dbgConsignments.Columns(22).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(22).Value = frmChildOptions.Code
    Case 23
        frmChildOptions.List = "Y or N"
        frmChildOptions.Code = dbgConsignments.Columns(23).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(23).Value = frmChildOptions.Code
    Case 24
        frmChildOptions.List = "Y or N"
        frmChildOptions.Code = dbgConsignments.Columns(24).Value
        frmChildOptions.Show vbModal
        dbgConsignments.Columns(24).Value = frmChildOptions.Code
    Case Else
       '
    End Select

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    On Error GoTo ErrHandler
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datConsignments.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datConsignments.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
   
    If gstrSystemRoute <> srCompanyRoute Then
        datConsignments.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    mstrLPTPortNumber = "LPT1"
    FillList "Consignment Status", cboConsignStatus, lstrConsignStatus()
    
    Select Case Me.Route
    Case gconstrConsignmentNorm
        datConsignments.RecordSource = "select * from PForce where CustNum=0"
            
        cboConsignStatus.Enabled = False
        cboConsignStatus.BackColor = vbActiveBorder
        txtSearchCriteria.Enabled = True
        txtSearchCriteria.BackColor = vbWindowBackground
        cmdPrintThermals.Enabled = False

    Case gconstrThermalPrintRun
        mbooDumpedSetup = False
        dbgConsignments.Top = fraSearchBy.Top
        dbgConsignments.Height = 5175
        datConsignments.RecordSource = "select * from PForce where Status = 'A'"
        datConsignments.Refresh
        cboConsignStatus.Enabled = False
        txtSearchCriteria.Enabled = False
        txtSearchCriteria.Visible = False
        cmdFind.Enabled = False
        cmdFind.Visible = False
        cmdRePrint.Enabled = False
        fraSearchBy.Enabled = False
        fraSearchBy.Visible = False
        cmdCreateFile.Enabled = False
        cmdPrintManifest.Enabled = False
        cmdBack.Caption = "&Cancel"
        
    End Select

    ShowBanner Me, Me.Route
    
    SetupHelpFileReqs
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "frmPForce.Load", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

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
        
    With cmdRePrint
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdCreateFile
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdRePrint.Width + cmdRePrint.Left + 120
    End With
    
    With cmdPrintManifest
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdCreateFile.Width + cmdCreateFile.Left + 120
    End With
    
    With cmdDumpSetup
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdPrintManifest.Width + cmdPrintManifest.Left + 120
    End With
    
    With cmdPrintThermals
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdDumpSetup.Width + cmdDumpSetup.Left + 120
    End With
    
    With dbgConsignments
        .Width = Me.Width - 360
        If (cmdHelp.Top - .Top) < 425 Then
            .Height = 425 - (cmdHelp.Top - .Top)
        Else
            .Height = (cmdHelp.Top - .Top) - 425
        End If
    End With
End Sub

Private Sub optSearchField_Click(Index As Integer)

    Select Case Index
    Case 4
        cboConsignStatus.Enabled = True
        cboConsignStatus.BackColor = vbWindowBackground
        txtSearchCriteria.Enabled = False
        txtSearchCriteria.BackColor = vbActiveBorder
    Case Else
        cboConsignStatus.Enabled = False
        cboConsignStatus.BackColor = vbActiveBorder
        txtSearchCriteria.Enabled = True
        txtSearchCriteria.BackColor = vbWindowBackground
    End Select
    
End Sub

Private Sub timActivity_Timer()
    
    CheckActivity
        
End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/PForce.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_PFORCE_MAIN
    ctlBanner1.WhatIsID = IDH_PFORCE_MAIN

    txtSearchCriteria.WhatsThisHelpID = IDH_PFORCE_SEATCRIT
    cmdFind.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    fraSearchBy.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(0).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(1).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(2).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(3).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(4).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    cboConsignStatus.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    dbgConsignments.WhatsThisHelpID = IDH_PFORCE_GRIDCONS
    cmdRePrint.WhatsThisHelpID = IDH_PFORCE_REPRINT
    cmdCreateFile.WhatsThisHelpID = IDH_PFORCE_CREATEEFILE
    cmdPrintManifest.WhatsThisHelpID = IDH_PFORCE_PRINTMANI
    cmdDumpSetup.WhatsThisHelpID = IDH_PFORCE_DUMP
    cmdPrintThermals.WhatsThisHelpID = IDH_PFORCE_PRINTTHERM
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
