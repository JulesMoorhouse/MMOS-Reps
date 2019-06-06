VERSION 5.00
Begin VB.Form frmSummary 
   Caption         =   "Form1"
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
      TabIndex        =   2
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.Frame fraKey 
      Caption         =   "Key"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   6360
      Width           =   2655
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&Info"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bad"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Good"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9120
      TabIndex        =   0
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   4
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin VB.Label lblSummaryItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Item"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Shape shpSummaryItem 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub cmdInfo_Click()

    MsgBox "When items appear in Red (Bad) or Green (Good) this relates to " & vbCrLf & _
           "comparitive calculations performed on other known statistics." & vbCrLf & vbCrLf & _
           "The status is assigned when in 10% of the comparitive statistic." & vbCrLf & vbCrLf & _
           "Please move your mouse over a statistic for more details." & vbCrLf & vbCrLf & _
           "The current statistics were last updated X", _
            vbInformation, gconstrTitlPrefix & "Summary Information"
End Sub

Private Sub Form_Activate()
Dim lintRetVal As Integer
Dim lstrSystemString As String
Dim lstrSQL As String
Dim lstrSummaryString As String
Dim lstrSummaryLastUpdated As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    PopulateSummaryArray
    lstrSystemString = PopNCalcSysRecStr
    lstrSQL = "UPDATE System SET System.Item = 'SysSummary', " & _
        "System.DateCreated = #" & Now() & "#, " & _
        "System.[Value] = '" & lstrSystemString & "'"
    gdatCentralDatabase.Execute lstrSQL

    ShowStatus 77
    
    lstrSummaryLastUpdated = GetSummaryString(lstrSummaryString)
    If lstrSummaryString <> "" Then
        AddValuestoSummaryArray lstrSummaryString & Chr(182)
        ShowSummaryInfo True, lstrSummaryLastUpdated
    End If
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If


    NameForm Me
        
    ShowBanner Me
    
    SetupHelpFileReqs
    
End Sub
Sub ShowSummaryInfo(pbooWithBuild As Boolean, Optional pstrLastUpdated As String)
Dim lintObjects As Integer
Dim lintArrInc As Integer
Dim llngLeftPos As Long
Const lconlngWidth = 4270

    If Me.Height < 7365 And Me.WindowState <> vbMaximized Then
        Me.Height = 7365
    End If
    llngLeftPos = 120
    
    If pbooWithBuild = True Then
        lintObjects = BuildSummaryItems(lblSummaryItem, shpSummaryItem, "CLIENT", 3800, Me)
    Else
        lintObjects = lblSummaryItem.Count
    End If
    
    For lintArrInc = 1 To lintObjects - 1
        shpSummaryItem(lintArrInc).Top = lblSummaryItem(lintArrInc - 1).Top + lblSummaryItem(lintArrInc - 1).Height + 40
        lblSummaryItem(lintArrInc).Top = shpSummaryItem(lintArrInc).Top + 25
                
        If shpSummaryItem(lintArrInc).Top + shpSummaryItem(lintArrInc).Height > (Me.Height - fraKey.Height) Then
            lblSummaryItem(lintArrInc).Top = shpSummaryItem(0).Top + 30 '+ 25
            shpSummaryItem(lintArrInc).Top = shpSummaryItem(0).Top + 5
            llngLeftPos = llngLeftPos + lconlngWidth
        End If
        
        lblSummaryItem(lintArrInc).Left = llngLeftPos + 120
        shpSummaryItem(lintArrInc).Left = llngLeftPos
        
    Next lintArrInc

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
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Summary.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_SUMRY_MAIN
    ctlBanner1.WhatIsID = IDH_SUMRY_MAIN

    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
