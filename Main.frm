VERSION 5.00
Begin VB.Form frmMainReps 
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFeatures 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1755
      Left            =   2985
      TabIndex        =   6
      Top             =   5580
      Width           =   7555
      Begin VB.ListBox lstNewFeatures 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   7
         Top             =   255
         Width           =   7550
      End
      Begin VB.CommandButton cmdFeatClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7290
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   20
         Width           =   260
      End
      Begin VB.CheckBox chkAllProgs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         Caption         =   "All programs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblNewFeatures 
         BackColor       =   &H80000002&
         Caption         =   " New Features: (Click for more information)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7555
      End
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   9480
      Top             =   3480
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1244
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1852
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   7335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks in advance!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mindwarp Consultancy Ltd.  March 2002."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":053F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please help us to help you!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If you have any comments or suggestions about this"
      Height          =   255
      Left            =   3060
      TabIndex        =   5
      Top             =   3600
      Width           =   7545
   End
   Begin VB.Label lblMCLContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "email@example.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4580
      MouseIcon       =   "Main.frx":0637
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Click me to make contact"
      Top             =   4320
      Width           =   4475
   End
   Begin VB.Label lblCover 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your cover has expired!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   7545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "program or its sister programs, please email :-"
      Height          =   255
      Left            =   3060
      TabIndex        =   2
      Top             =   3960
      Width           =   7545
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      X1              =   1560
      X2              =   2160
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      X1              =   720
      X2              =   1560
      Y1              =   2160
      Y2              =   1440
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      X1              =   720
      X2              =   1320
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      X1              =   1320
      X2              =   2160
      Y1              =   2520
      Y2              =   1800
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      X1              =   1320
      X2              =   2160
      Y1              =   3000
      Y2              =   2280
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      X1              =   720
      X2              =   1320
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      X1              =   720
      X2              =   960
      Y1              =   2640
      Y2              =   2400
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      X1              =   1920
      X2              =   2160
      Y1              =   2160
      Y2              =   2280
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      X1              =   1320
      X2              =   2160
      Y1              =   3480
      Y2              =   2760
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      X1              =   720
      X2              =   1320
      Y1              =   3120
      Y2              =   3480
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      X1              =   720
      X2              =   960
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      X1              =   1920
      X2              =   2160
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Shape shpBacking 
      BorderColor     =   &H00800000&
      BorderWidth     =   10
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   6195
      Left            =   100
      Top             =   1140
      Width           =   2775
   End
End
Attribute VB_Name = "frmMainReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lvarOrigBackcolor As Variant
Dim llngNewFeatures() As Long

Private Sub cmdFeatClose_Click()

    fraFeatures.Visible = False
    
End Sub

Private Sub Form_Activate()

    If gstrTempKeyFail <> "" Then
        MsgBox "Please be advised that your temporary license will expire on " & gstrTempKeyFail & " after this date, " & vbCrLf & _
            "this software will no longer function! For continued usage beyond this date please " & vbCrLf & _
            "ensure that you have purchased a full license. " & vbCrLf & vbCrLf & _
            "If you would like to discuss this matter, please Contact Mindwarp Consultancy Ltd.", vbInformation, gconstrTitlPrefix & "Warning!"
        gstrTempKeyFail = ""
    End If
    
End Sub
Private Sub Form_Load()
Dim lstrShowFeatures As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    lvarOrigBackcolor = lblMCLContact.ForeColor
    
    ShowBanner Me
    gstrButtonRoute = gconstrMainMenu
    mdiMain.DrawButtonSet gstrButtonRoute
        
    lstrShowFeatures = GetSetting(gstrIniAppName, "UI", "ShowFeatures")
    
    lblMCLContact.Caption = gstrOurContactWeb

    If IsBlank(lstrShowFeatures) Then
        lstrShowFeatures = True
    ElseIf UCase$(lstrShowFeatures) <> "TRUE" Or UCase$(lstrShowFeatures) <> "FALSE" Then
        lstrShowFeatures = True
    End If
    
    If CBool(lstrShowFeatures) = True Then
        fraFeatures.Visible = True
    Else
        lstrShowFeatures = False
        fraFeatures.Visible = False
        SaveSetting gstrIniAppName, "UI", "ShowFeatures", lstrShowFeatures
    End If
    
    PopFeatList lstNewFeatures, False, llngNewFeatures()
    
    If gdatCoverDate < Date And gdatCoverDate <> "00:00:00" Then
        lblCover = "Your cover has expired! " & gdatCoverDate & " " & gstrStatic.strUnlockCode
        lblCover.Visible = True
    End If
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMCLContact.ForeColor = lvarOrigBackcolor
    
End Sub
Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    With lblCover
        .Width = Me.Width - 3060
    End With
    
    With Label2
        .Width = Me.Width - 3060
    End With
    
    With Label3
        .Width = Me.Width - 3060
    End With
            
    With shpBacking
        .Left = Me.Left + 160
        .Height = (Me.Height - (705 + 1080)) - 180
    End With
    
    With fraFeatures
        .Top = ((Me.Height - 705) - 1755) - 100
        .Left = shpBacking.Width + shpBacking.Left + 70
        .Width = (Me.Width - .Left) - 110
    End With
    
    With lblMCLContact
        .Left = fraFeatures.Left + ((fraFeatures.Width / 2) - (.Width / 2))
    End With
    
    lstNewFeatures.Width = fraFeatures.Width
    lblNewFeatures.Width = fraFeatures.Width
    chkAllProgs.Left = fraFeatures.Width - 1795
    cmdFeatClose.Left = fraFeatures.Width - 265
    lblNewFeatures.BackColor = vbActiveTitleBar
    chkAllProgs.BackColor = vbActiveTitleBar
    
    shpBacking.Visible = False
    shpBacking.Visible = True
    
End Sub
Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub lblMCLContact_Click()

   Dim StartDoc As Long

     StartDoc = ShellExecute(Me.hwnd, "open", gstrOurContactWeb, _
       "", "C:\", 1)

End Sub

Private Sub lblMCLContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMCLContact.ForeColor = vbRed
    
End Sub
Private Sub lstNewFeatures_Click()

    FeatureMsg (llngNewFeatures(lstNewFeatures.ListIndex))
    
End Sub
Private Sub chkAllProgs_Click()

    If chkAllProgs.Value = 0 Then ' unchecked
        PopFeatList lstNewFeatures, False, llngNewFeatures()
    Else
        PopFeatList lstNewFeatures, True, llngNewFeatures()
    End If
    
End Sub
