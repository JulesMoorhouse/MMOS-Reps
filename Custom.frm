VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Reports"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Preview"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdMaxMin 
      Caption         =   "&Maximize"
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtThisQuery 
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdEndDate 
      Caption         =   "&Get End Date"
      Height          =   492
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton cmdStartDate 
      Caption         =   "&Get Start Date"
      Height          =   492
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox txtSQL 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboCustomRepSelect 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
   Begin VB.Data datCustom 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   492
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5670
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "08/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:45"
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid dbgCustomGrid 
      Bindings        =   "Custom.frx":0000
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "Custom.frx":0018
      TabIndex        =   3
      Top             =   1320
      Width           =   9975
   End
   Begin VB.Label lblFoundNumber 
      Caption         =   "Found 0 recrods"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblEndDate 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblStartDate 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Report name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrSysDB As String

Function ChkNull(pvar As Variant) As String

    If IsNull(pvar) Then
        ChkNull = ""
    Else
        ChkNull = pvar
    End If
    
End Function

Private Sub cboCustomRepSelect_Click()
Dim lbooLocked As Boolean
Dim llngSeqNum As Long
Dim lintInUse As Integer

    'does it matter whether it's local or central
    'I guess it does as this is the only place where datCustom knows
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        If lstrSysDB = "CENTRAL" Then
            datCustom.DatabaseName = gstrStatic.strCentralTestingDBFile
        Else
            datCustom.DatabaseName = gstrStatic.strLocalTestingDBFile
        End If
    Case gconstrLiveMode
        If lstrSysDB = "CENTRAL" Then
            datCustom.DatabaseName = gstrStatic.strCentralDBFile
        Else
            datCustom.DatabaseName = gstrStatic.strLocalDBFile
        End If
    End Select
        
    If gstrSystemRoute <> srCompanyRoute Then
        If lstrSysDB = "CENTRAL" Then
            datCustom.Connect = gstrDBPasswords.strCentralDBPasswordString
        Else
            datCustom.Connect = gstrDBPasswords.strLocalDBPasswordString
        End If
    End If
    
    datCustom.RecordSource = ""
    
End Sub
Function DateOK(pstrParam As String) As Boolean

    DateOK = True
    Select Case pstrParam
    Case "START"
        If Not IsDate(lblStartDate) Then
            MsgBox "The Start date is not Valid!", , gconstrTitlPrefix & "Mandatory Field"
            DateOK = False
        End If
    Case "END"
        If Not IsDate(lblEndDate) Then
            MsgBox "The End date is not Valid!", , gconstrTitlPrefix & "Mandatory Field"
            DateOK = False
        End If
    Case "S&E"
        If Not IsDate(lblStartDate) Then
            MsgBox "The Start date is not Valid!", , gconstrTitlPrefix & "Mandatory Field"
            DateOK = False
        End If
        If Not IsDate(lblEndDate) Then
            MsgBox "The End date is not Valid!", , gconstrTitlPrefix & "Mandatory Field"
            DateOK = False
        End If
    End Select
    
End Function

Private Sub cmdDisplay_Click()
Dim lintArrInc As Integer

    If Not DateOK("S&E") Then Exit Sub
    
    dbgCustomGrid.ClearFields

    ReplaceParams txtSQL, txtThisQuery, lblStartDate, lblEndDate

    datCustom.RecordSource = txtThisQuery
    datCustom.Refresh
    dbgCustomGrid.Caption = cboCustomRepSelect
    lblFoundNumber = "Found " & datCustom.Recordset.RecordCount & " records."

    gstrReport.strDelimDetailsFile = GetTempDir & "D" & Format(Now(), "MMDDSSN") & ".tmp"
   
    AnalyseFields datCustom.RecordSource, lstrSysDB, frmPrintPreview.picCanvas
    AnalyseSQL datCustom.RecordSource, lstrSysDB, frmPrintPreview.picCanvas

    For lintArrInc = 0 To UBound(lstrFieldNames)
        
        dbgCustomGrid.Columns(lintArrInc).Width = lstrFieldNames(lintArrInc).lngFieldLength + 100
    Next lintArrInc
   
    ClearReportBuffer
    
End Sub

Private Sub cmdEndDate_Click()

    lblEndDate = CheckCalendar(vbKeyInsert, lblEndDate)

End Sub

Private Sub cmdMaxMin_Click()

    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
        cmdMaxMin.Caption = "&Maximize"
    Else
        Me.WindowState = vbMaximized
        cmdMaxMin.Caption = "&Normal"
    End If
    
End Sub

Private Sub cmdPrint_Click()
Dim lintArrInc As Integer

    If datCustom.RecordSource = "" Then
        MsgBox "You must first select a report and click display!", vbInformation, gconstrTitlPrefix & "Print"
        Exit Sub
    End If
    
    gintScaleFactor = 1
    gbooTotalLineRequired = False
    
    Font.Name = "Arial"
    Font.Size = 11 / gintScaleFactor
   
    AnalyseFields datCustom.RecordSource, lstrSysDB, frmPrintPreview.picCanvas
        
    With gstrReport
        .strReportName = cboCustomRepSelect
        .strStartRangeDate = lblStartDate
        .strEndRangeDate = lblEndDate
        .booBarsOn = True
        .intSpacing = rpSpacing.rpSpacingSingle
        .lngMargins = rpMargins.rpMarginNarrow
        .sngFontSize = rpFontFactor.rpFontFactorNormal
        .strDelimDetailsFile = GetTempDir & "D" & Format(Now(), "MMDDSSN") & ".tmp"
       
        AnalyseSQL datCustom.RecordSource, lstrSysDB, frmPrintPreview.picCanvas
        
        frmPrintPreview.Show vbModal
        Set frmPrintPreview = Nothing
        Kill .strDelimDetailsFile
    End With
    
    ClearReportBuffer
    
End Sub

Private Sub cmdStartDate_Click()

    lblStartDate = CheckCalendar(vbKeyInsert, lblStartDate)

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datCustom.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datCustom.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
   
    If gstrSystemRoute <> srCompanyRoute Then
        datCustom.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
   
    dbgCustomGrid.ClearFields
    
    If cboCustomRepSelect.ListCount >= 0 Then
        cboCustomRepSelect.ListIndex = 0
    End If
    
    lblEndDate = Format$(Date, "dd/mmm/yyyy")
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()

    With cmdClose
        .Top = Me.Height - 1275
        .Left = Me.Width - 1545
    End With
    
    With dbgCustomGrid
        .Height = Me.Height - 2700
        .Width = Me.Width - 330
    End With
    
End Sub
Function CashOrCheque(pstrPaytype1 As String, pstrPaytype2 As String, _
    pcurPay1 As Currency, pcurPay2 As Currency, pstrTypeWanted As String) As Currency
Dim lcurCash As Currency
Dim lcurCreditCard As Currency
    
    Select Case Trim$(pstrPaytype1)
    Case "Q", "V" 'Cheque & Voucher
        lcurCash = pcurPay1
    Case "C" 'Credit Card
        lcurCreditCard = pcurPay1
    End Select
    
    Select Case Trim$(pstrPaytype2)
    Case "Q", "V" 'Cheque & Voucher
        lcurCash = pcurPay2
    Case "C" 'Credit Card
        lcurCreditCard = pcurPay2
    End Select
    
    Select Case Trim$(pstrTypeWanted)
    Case "C"
        CashOrCheque = lcurCreditCard
    Case "Q"
        CashOrCheque = lcurCash
    End Select

End Function
Function ChequeOrCreditCount(pstrPaytype1 As String, pstrPaytype2 As String, _
    pcurPay1 As Currency, pcurPay2 As Currency, pstrTypeWanted As String) As Integer
Dim lintCreditCard As Integer
Dim lintCheque As Integer

    Select Case Trim$(pstrPaytype1)
    Case "Q", "V" 'Cheque & Voucher
        lintCheque = 1
        lintCreditCard = 0
    Case "C" 'Credit Card
        lintCheque = 0
        lintCreditCard = 1
    End Select
    
    Select Case Trim$(pstrPaytype2)
    Case "Q", "V" 'Cheque & Voucher
        lintCheque = 1
        lintCreditCard = 0
    Case "C" 'Credit Card
        lintCheque = 0
        lintCreditCard = 1
    End Select
    
    Select Case Trim$(pstrTypeWanted)
    Case "C"
        ChequeOrCreditCount = lintCreditCard
    Case "Q"
        ChequeOrCreditCount = lintCheque
    End Select

End Function
