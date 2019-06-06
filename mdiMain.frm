VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "Mindwarp Mail Order System"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   3960
   End
   Begin VB.PictureBox picListBar 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   8040
      Left            =   0
      ScaleHeight     =   7980
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "4/16/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:27 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&1 "
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder1 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory1 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder1 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&2"
         Index           =   1
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder2 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory2 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder2 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&3"
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder3 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory3 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder3 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&4"
         Index           =   3
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder4 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory4 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder4 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistory 
         Caption         =   "&5"
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu mnuFileHistoryModOrder5 
            Caption         =   "&Modify This Order"
         End
         Begin VB.Menu mnuFileHistoryOrdHistory5 
            Caption         =   "Orders &History"
         End
         Begin VB.Menu mnuFileHistoryPackOrder5 
            Caption         =   "&Pack This Order"
         End
      End
      Begin VB.Menu mnuFileHistorySep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewShowPicBar 
         Caption         =   "Show &Picture Bar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewShowNewFeatures 
         Caption         =   "Show New &Features"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMaxOnStartup 
         Caption         =   "&Maximize On Startup"
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
      Begin VB.Menu mnuGoItem1 
         Caption         =   "Item1"
      End
      Begin VB.Menu mnuGoItem2 
         Caption         =   "Item2"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem3 
         Caption         =   "Item3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem4 
         Caption         =   "Item4"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem5 
         Caption         =   "Item5"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoItem6 
         Caption         =   "Item6"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsMinder 
         Caption         =   "&Minder Full"
      End
      Begin VB.Menu mnuToolsResetGrid 
         Caption         =   "Reset &Grid(s) Layout"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsConfigureValues 
         Caption         =   "&Configure Values"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsMaintainProducts 
         Caption         =   "Maintain &Products"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsEssentialSettings 
         Caption         =   "Essential &Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsChangePassword 
         Caption         =   "Change Pass&word"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExternalPrograms 
         Caption         =   "&External Programs"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents and Index	F1"
      End
      Begin VB.Menu mnuHelpWhatsThis 
         Caption         =   "What's This?	Shift + F1"
      End
      Begin VB.Menu mnuHelpTutorial 
         Caption         =   "&Tutorial"
      End
      Begin VB.Menu mnuHelpQuickStart 
         Caption         =   "&Quick Start Sheets"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCFU 
         Caption         =   "Check For &Updates"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lintCurrOrderEntryButton As Integer
Dim lintCurrOrderEnqButton As Integer
Dim lintCurrAcctMaintButton As Integer
Dim lintCurrFinanceButton As Integer
Dim lintCurrPackingButton As Integer
Dim lintCurrOrderMaintButton As Integer

Private Sub MDIForm_Activate()

    sbStatusBar.Panels(2).Text = gstrGenSysInfo.strUserName
    
End Sub

Private Sub MDIForm_Load()
        
    MDILoad Me, frmMainReps
    
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lintRetVal As Integer
Dim lstrExitMsg As String

    If gintForceAppClose = fcCompleteClose Or gintForceAppClose = fcCloseKeepDB Then
       
    Else
         Select Case gstrButtonRoute
         Case gconstrEntry, gconstrEnquiry, gconstrAccount
             If gstrCurrentLoadedForm.Name <> "frmCustAcctSel" Then
                 lstrExitMsg = "WARNING: closing the system from this screen may result" & vbCrLf & _
                     "in information being lost!" & vbCrLf & vbCrLf
             End If
         
         End Select
        
         lintRetVal = MsgBox(lstrExitMsg & "You are about to logout and close the system! Procced?", _
             vbYesNo + vbDefaultButton1 + vbExclamation, gconstrTitlPrefix & "System Exit")
         
         If lintRetVal = vbNo Then
             Cancel = True
             Exit Sub
         End If
    End If
    
    ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    
    If gintForceAppClose <> fcCloseKeepDB Then
        Busy True, Me
        gdatCentralDatabase.Close
        gdatLocalDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
    End If
    
    If UCase$(App.ProductName) <> "LITE" Then
        UpdateLoader
    End If
    
    Busy False, Me
    
    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
    Unload Me
    
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState = vbNormal Then
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
    End If

End Sub

Private Sub MDIForm_Terminate()

    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Unload frmButtons
    If Not DebugVersion Then
        'Stop subclassing.
        Unhook
    End If
    
End Sub

Private Sub mnuHelpContents_Click()

    StandardMenuOptions mnuHelpContents.Caption
    
End Sub

Private Sub mnuHelpQuickStart_Click()

    StandardMenuOptions mnuHelpQuickStart.Caption
    
End Sub

Private Sub mnuHelpTutorial_Click()

    StandardMenuOptions mnuHelpTutorial.Caption
    
End Sub

Private Sub mnuHelpWhatsThis_Click()

    StandardMenuOptions mnuHelpWhatsThis.Caption
    
End Sub

Private Sub picListBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicListBarMouseDown Me, Button, Shift, X, Y
    
End Sub
Private Sub picListBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PicListBarMouseMove Me, Button, Shift, X, Y

End Sub
Private Sub picListBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    gbooUIScrollButtonClicked = False
    
End Sub
Sub ButtonSelected(pintButtonIndex As Integer)
Dim lintRetVal As Integer

    Select Case gstrButtonRoute
    Case gconstrMainMenu
        Select Case pintButtonIndex
        Case 0
            UnloadLastForm
            gstrButtonRoute = gconstrGenralReporting
            Set gstrCurrentLoadedForm = frmReports
            DrawButtonSet gstrButtonRoute
            frmReports.Show
        Case 1
            UnloadLastForm
            gstrButtonRoute = gconstrConsignmentNorm
            Set gstrCurrentLoadedForm = frmPForce
            frmPForce.Route = gconstrConsignmentNorm
            frmPForce.CallingForm = frmMainReps
            DrawButtonSet gstrButtonRoute
            frmPForce.Show
        Case 2
            UnloadLastForm
            gstrButtonRoute = gconstrMarketingData
            FillSystemListsDefs
            FillListsMultiDef
            Set gstrCurrentLoadedForm = frmStaMultiMarket
            frmStaMultiMarket.Route = gconstrAdminRoute
            frmStaMultiMarket.CallingForm = frmMainReps
            DrawButtonSet gstrButtonRoute
            frmStaMultiMarket.Show
        Case 3
            MsgBox "Agent reporting will provide information about users, e.g. audit trail, orders taken " & vbCrLf & _
                "and money made." & vbCrLf & vbCrLf & gstrComingSoon, vbInformation, gconstrTitlPrefix & "Coming Soon!"
            frmMainReps.Show
        Case 4
            UnloadLastForm
            gstrButtonRoute = gconstrSummaryInfo
            Set gstrCurrentLoadedForm = frmSummary
            DrawButtonSet gstrButtonRoute
            frmSummary.Show
        Case 5
            UnloadLastForm
            gstrButtonRoute = gconstrDuplicateHandling
            Set gstrCurrentLoadedForm = frmDuplicates
            DrawButtonSet gstrButtonRoute
            frmDuplicates.Show
        End Select
    Case gconstrGenralReporting
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmMainReps
            Call frmReports.cmdBack_Click
        End Select
    Case gconstrConsignmentNorm
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            'loose changes and return to menu
            'Clear functions here
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmMainReps
            Call frmPForce.cmdBack_Click
        End Select
    Case gconstrThermalPrintRun
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmReports
            Call frmPForce.cmdBack_Click
        End Select
    Case gconstrMarketingData
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmMainReps
            Call frmStaMultiMarket.cmdProceed_Click
        End Select
    Case gconstrSummaryInfo
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmMainReps
            Call frmSummary.cmdBack_Click
        End Select
    Case gconstrDuplicateHandling
        Select Case pintButtonIndex
        Case 0
            'Do Nothing
        Case 1 'return to menu
            gstrButtonRoute = gconstrMainMenu
            UnloadLastForm
            Set gstrCurrentLoadedForm = frmDuplicates
            Call frmDuplicates.cmdBack_Click
        End Select
    End Select
    
End Sub
Sub DrawButtonSet(pstrRoute As String, Optional pstrParam As Variant)
Dim llngDownVar As Long

    lintCurrOrderEntryButton = -1
    lintCurrOrderEnqButton = -1
    lintCurrAcctMaintButton = -1
    lintCurrFinanceButton = -1
    lintCurrPackingButton = -1
    lintCurrOrderMaintButton = -1
    
    picListBar.Cls
    
    If IsMissing(pstrParam) Then pstrParam = ""
    
    gstrButtonRoute = pstrRoute
    
    If gstrUILastButtonRoute <> pstrRoute And gstrUILastButtonRoute <> "" Then
        gconUITopPos = gconUIButtonTopPosDefault
    End If
    
    gstrUILastButtonRoute = pstrRoute
    Select Case pstrRoute
    Case gconstrMainMenu
            DrawButton Me, 0, 0, 16, "General", "Reporting"
            DrawButton Me, 1, 0, 6, "Distribution" ':             lintCurrDistributionButton = 4
            DrawButton Me, 2, 0, 14, "Marketing", "Settings"
            DrawButton Me, 3, 0, 21, "Agent", "Reporting"
            DrawButton Me, 4, 0, 15, "Summary", "Info"
            DrawButton Me, 5, 0, 17, "Duplicate", "Handling"
            gintUINumberofButtonsDraw = 5
    Case gconstrGenralReporting
        DrawButton Me, 0, 0, 16, "General", "Reporting"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrConsignmentNorm
        DrawButton Me, 0, 0, 6, "Distribution"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrThermalPrintRun
        DrawButton Me, 0, 0, 6, "Thermal Label", "Printing"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrMarketingData
        DrawButton Me, 0, 0, 14, "Marketing", "Settings"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrAgentReporting
    Case gconstrSummaryInfo
        DrawButton Me, 0, 0, 15, "Summary", "Info"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    Case gconstrDuplicateHandling   
        DrawButton Me, 0, 0, 17, "Duplicate", "Handling"
        DrawButton Me, 1, 0, 9, "Back"
        gintUINumberofButtonsDraw = 1
    End Select
    
    FinishDrawingButtonSet Me, llngDownVar, pstrParam

End Sub

Private Sub picListBar_Resize()

    gconUITopPos = gconUIButtonTopPosDefault
    DrawButtonSet gstrButtonRoute

End Sub

Private Sub Timer1_Timer()

    CheckActivity

End Sub

Private Sub mnuEditCopy_Click()

    StandardMenuOptions mnuEditCopy.Caption

End Sub

Private Sub mnuEditCut_Click()

    StandardMenuOptions mnuEditCut.Caption

End Sub

Private Sub mnuEditPaste_Click()
    
    StandardMenuOptions mnuEditPaste.Caption

End Sub

Private Sub mnuFileExit_Click()

    StandardMenuOptions mnuFileExit.Caption
    
End Sub

Private Sub mnuFilePrintSetup_Click()

    StandardMenuOptions mnuFilePrintSetup.Caption
    
End Sub

Private Sub mnuGoItem1_Click()

    MenuCommands mnuGoItem1.Caption
    
End Sub

Private Sub mnuGoItem2_Click()

    MenuCommands mnuGoItem2.Caption
    
End Sub

Private Sub mnuGoItem3_Click()

    MenuCommands mnuGoItem3.Caption
    
End Sub

Private Sub mnuGoItem4_Click()

    MenuCommands mnuGoItem4.Caption
    
End Sub

Private Sub mnuGoItem5_Click()

    MenuCommands mnuGoItem5.Caption
    
End Sub

Private Sub mnuGoItem6_Click()

    MenuCommands mnuGoItem6.Caption
    
End Sub

Private Sub mnuHelpAbout_Click()

    StandardMenuOptions mnuHelpAbout.Caption
    
End Sub

Private Sub mnuHelpCFU_Click()

    MenuCommands mnuHelpCFU.Caption
    
End Sub

Private Sub mnuToolsChangePassword_Click()

    MenuCommands mnuToolsChangePassword.Caption
    
End Sub

Private Sub mnuToolsConfigureValues_Click()

    MenuCommands mnuToolsConfigureValues.Caption
    
End Sub

Private Sub mnuToolsEssentialSettings_Click()

    MenuCommands mnuToolsEssentialSettings.Caption
    
End Sub

Private Sub mnuToolsExternalPrograms_Click()

    MenuCommands mnuToolsExternalPrograms.Caption
    
End Sub

Private Sub mnuToolsMaintainProducts_Click()

    MenuCommands mnuToolsMaintainProducts
    
End Sub

Private Sub mnuToolsMinder_Click()

    StandardMenuOptions mnuToolsMinder.Caption
    
End Sub

Private Sub mnuToolsResetGrid_Click()

    StandardMenuOptions mnuToolsResetGrid.Caption
    
End Sub

Private Sub mnuViewMaxOnStartup_Click()

    StandardMenuOptions mnuViewMaxOnStartup.Caption
    
End Sub

Private Sub mnuViewShowNewFeatures_Click()

    StandardMenuOptions mnuViewShowNewFeatures.Caption
    
End Sub

Private Sub mnuViewShowPicBar_Click()

    StandardMenuOptions mnuViewShowPicBar.Caption
    
End Sub
Sub MenuCommands(pstrItem As String)

    Select Case pstrItem
    Case mnuMgrGoGenReps
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 0
    Case mnuMgrGoDistribution
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 1
    Case mnuMgrGoMarkSets
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 2
    Case mnuMgrGoAgentReps
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 3
    Case mnuMgrGoSumInfo
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 4
    Case mnuMgrGoDupHand
        gstrButtonRoute = gconstrMainMenu
        ButtonSelected 5
    Case mnuToolsChangePassword.Caption
        frmChildUserPass.Route = "PASSCHANGE"
        frmChildUserPass.Show vbModal
    End Select
    
End Sub
