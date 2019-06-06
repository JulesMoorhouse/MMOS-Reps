VERSION 5.00
Begin VB.Form frmChildLabelOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Label Options"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtLeftMargin 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "2"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtTopMargin 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtLinesBetweenLabels 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "5"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtCharsLeftToRight 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "35"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Preview"
      Height          =   360
      Left            =   3720
      TabIndex        =   7
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      TabIndex        =   9
      Top             =   4080
      Width           =   1305
   End
   Begin VB.ComboBox cboLabelType 
      Height          =   315
      ItemData        =   "LabelOps.frx":0000
      Left            =   2040
      List            =   "LabelOps.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtLabelsDownPage 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "7"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtLabelsAcrossPage 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "3"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ensure that settings below allow all your labels to fit on each page!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Left margin (in chars)"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Top margin (in chars)"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Lines between labels:"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Chars between labels (Left to Right)"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label type :"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Labels down  page:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Labels across page:"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmChildLabelOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrLabelLaout() As String
Dim lstrLabelFont() As String
Dim lstrLabelNumbers() As String

Dim mbooCancelled As Boolean

Dim mobjPrintingObject As Object
Dim lstrScreenHelpFile As String

Public Property Let PrintingObject(pstrPrintingObject As Object)

    Set mobjPrintingObject = pstrPrintingObject

End Property
Public Property Get PrintingObject() As Object

    PrintingObject = mobjPrintingObject
    
End Property
Public Property Let Cancelled(pstrCancelled As Boolean)

    mbooCancelled = pstrCancelled

End Property
Public Property Get Cancelled() As Boolean

    Cancelled = mbooCancelled
    
End Property

Private Sub cboLabelType_Click()
'0=Across, 1=Down, 2=VGap, 3=HGap, 4=TopMarg, 5=Leftmarg

    On Error Resume Next
    mobjPrintingObject.Font = ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelFont)), 1, ",")
    mobjPrintingObject.FontSize = ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelFont)), 2, ",")
    
    txtLabelsAcrossPage = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 1, ","))
    txtLabelsDownPage = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 2, ","))
    txtLinesBetweenLabels = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 3, ","))
    txtCharsLeftToRight = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 4, ","))
    txtTopMargin = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 5, ","))
    txtLeftmargin = Val(ReturnNthStr(Trim$(NotNull(cboLabelType, lstrLabelNumbers)), 6, ","))

End Sub

Private Sub cmdCancel_Click()

    mbooCancelled = True
    Unload Me
    
End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdPrint_Click()
'0=Across, 1=Down, 2=VGap, 3=HGap, 4=TopMarg, 5=Leftmarg

    mobjPrintingObject.Font = "Courier"
    mobjPrintingObject.FontSize = 12
    
    With gstrLabelPage
        .intLabelsAcross = Val(txtLabelsAcrossPage)
        .intLabelsDown = Val(txtLabelsDownPage)
        .lngVertGap = mobjPrintingObject.TextHeight("WWW") * Val(txtLinesBetweenLabels) '150
        .lngHorizGap = mobjPrintingObject.TextWidth(String(Val(txtCharsLeftToRight), "L")) '+ 150 '3000
        .lngTopMargin = mobjPrintingObject.TextHeight("W") * Val(txtTopMargin)
        .lngLeftMargin = mobjPrintingObject.TextWidth("W") * Val(txtLeftmargin)
    End With

    mbooCancelled = False
    Unload Me
    
End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    FillList "Label Layouts", cboLabelType, lstrLabelLaout(), lstrLabelFont(), lstrLabelNumbers()
    cboLabelType.ListIndex = 0
    mbooCancelled = True
    
    SetupHelpFileReqs
    
End Sub

Private Sub txtCharsLeftToRight_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtCharsLeftToRight_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtCharsLeftToRight_LostFocus()

    If Not IsNumeric(txtCharsLeftToRight) Then
        txtCharsLeftToRight = 0
    End If
    txtCharsLeftToRight = Trim$(txtCharsLeftToRight)
    
End Sub

Private Sub txtLabelsAcrossPage_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLabelsAcrossPage_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLabelsAcrossPage_LostFocus()

    If Not IsNumeric(txtLabelsAcrossPage) Then
        txtLabelsAcrossPage = 0
    End If
    txtLabelsAcrossPage = Trim$(txtLabelsAcrossPage)
    
End Sub

Private Sub txtLabelsDownPage_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLabelsDownPage_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLabelsDownPage_LostFocus()

    If Not IsNumeric(txtLabelsDownPage) Then
        txtLabelsDownPage = 0
    End If
    txtLabelsDownPage = Trim$(txtLabelsDownPage)
    
End Sub

Private Sub txtLeftMargin_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLeftMargin_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLeftMargin_LostFocus()

    If Not IsNumeric(txtLeftmargin) Then
        txtLeftmargin = 0
    End If
    txtLeftmargin = Trim$(txtLeftmargin)
    
End Sub

Private Sub txtLinesBetweenLabels_GotFocus()

    SetSelected Me

End Sub

Private Sub txtLinesBetweenLabels_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtLinesBetweenLabels_LostFocus()

    If Not IsNumeric(txtLinesBetweenLabels) Then
        txtLinesBetweenLabels = 0
    End If
    txtLinesBetweenLabels = Trim$(txtLinesBetweenLabels)
    
End Sub

Private Sub txtTopMargin_GotFocus()

    SetSelected Me

End Sub

Private Sub txtTopMargin_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtTopMargin_LostFocus()

    If Not IsNumeric(txtTopMargin) Then
        txtTopMargin = 0
    End If
    txtTopMargin = Trim$(txtTopMargin)
    
End Sub

Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/ChLabLay.xml>WhatsScreen"

    cboLabelType.WhatsThisHelpID = IDH_CHILABOPS_LABTYPE
    txtLabelsAcrossPage.WhatsThisHelpID = IDH_CHILABOPS_LABSACROSS
    txtLabelsDownPage.WhatsThisHelpID = IDH_CHILABOPS_LADSDOWN
    txtLinesBetweenLabels.WhatsThisHelpID = IDH_CHILABOPS_LINESBETW
    txtCharsLeftToRight.WhatsThisHelpID = IDH_CHILABOPS_CHARSLEFTORI
    txtTopMargin.WhatsThisHelpID = IDH_CHILABOPS_TOPMARG
    txtLeftmargin.WhatsThisHelpID = IDH_CHILABOPS_LEFTMARG
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    
End Sub
