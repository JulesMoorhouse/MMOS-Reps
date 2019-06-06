VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReports 
   Caption         =   "Please select a report ..."
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ControlBox      =   0   'False
   Icon            =   "Reports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtGroupings 
      Height          =   375
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "Reports.frx":014A
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrinterSetup 
      Caption         =   "&Page Setup"
      Height          =   360
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1215
      Left            =   9120
      TabIndex        =   21
      Top             =   3360
      Width           =   1345
      Begin VB.OptionButton optOutput 
         Caption         =   "&Graph"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "&Grid"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "&Spreadsheet"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1195
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "&Preview"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtThisQuery 
      Height          =   375
      Left            =   9600
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSQL 
      Height          =   375
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Reports.frx":0150
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdExternal 
      Caption         =   "     &External      (MS Access)"
      Height          =   480
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame fraOrderRange 
      Caption         =   "Order Range Selection"
      Height          =   1335
      Left            =   5400
      TabIndex        =   15
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox txtEndOrderNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEndOrderNum 
         Caption         =   "End Ord Num"
         Height          =   360
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1305
      End
      Begin VB.CommandButton cmdStartOrderNum 
         Caption         =   "Start Ord Num"
         Height          =   360
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1305
      End
      Begin VB.TextBox txtStartOrderNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCustomerEnd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblCustomerStart 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame fraDates 
      Caption         =   "Date Range Selection"
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   5175
      Begin VB.CommandButton cmdEndDate 
         Caption         =   "&Get End Date"
         Height          =   360
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1305
      End
      Begin VB.CommandButton cmdStartDate 
         Caption         =   "&Get Start Date"
         Height          =   360
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblEndDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblStartDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   9000
      Top             =   6000
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   9120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1305
   End
   Begin VB.ListBox lstReports 
      Height          =   3540
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   8580
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4290
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   7567
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Customer Services"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Finance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Packing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Discrepancies"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   27
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Most reports will ask for Start and End Dates, however some may not use them."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   6840
      Width           =   10575
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lcon2ndColPos = 5500

Dim lstrCreditCardClaim As String:
Dim lstrChequeRefundAdviceNotes As String
Dim lstrRefundCheques As String

Dim lstrReportSalesSummary As String

Dim lstrAdviceNoteSpecific As String:       Dim lstrAdviceNotes As String
Dim lstrAdviceNotesRange As String:         Dim lstrAdviceNotesPrinted As String
Dim lstrBatchPickings As String:            Dim lstrThermalLabels As String

Dim lstrReportStockAllocatedSummary As String: Dim lstrReportStockSummary As String

Dim lstrPrintObjAdviceNotes As String:     Dim lstrPrintObjRefundAdviceNotes As String
Dim lstrPrintObjPFManifest As String:      Dim lstrPrintObjRefundCheques As String
Dim lstrPrintObjBatchPick As String:      Dim lstrPrintObjCreditCardClaims As String
Dim lstrPrintObjCreditCardRefunds As String

Dim lstrPrintObjAdviceNoteSpecific As String
Dim lstrPrintObjAdviceNotesRange As String:         Dim lstrPrintObjAdviceNotesPrinted As String
Dim lstrPrintObjAdviceNoteStatusSpecific As String

Const gconstrReportTypeCustomerNameBlank = "Customer Name"
Dim lstrRepId() As String
Dim lstrSysDB As String
Dim lintOrientation As Integer
Dim lstrScreenHelpFile As String

Function FillReplaceParams(ByVal pobjTextBoxThisQuery As Object, pstrReportType As String) As Boolean
Dim llngPos As Long

    With frmChildGenericDropdown
    
        If InStr(1, pobjTextBoxThisQuery, "[Stock Batch Num]") > 0 Then
            GetCurrentStockBatchNumber
            .LabelStr = gdatLastStockBatchNumberDate & " (M" & glngStockBatchNumber & ")"
            .FormCaption = "Stock Batch Number Selection"
            .LabelCaption = "Date of last batch:"
            .CodeField = "StockBatchNum"
            .DescField = "StockBatchNum"
            .SQL = "SELECT DISTINCT StockBatchNum From AdviceNotes Where (((StockBatchNum) Is Not Null)) ORDER BY AdviceNotes.StockBatchNum DESC;"
            .Show vbModal
            If .Cancelled = True Then
                FillReplaceParams = True
                Exit Function
            End If
            pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[Stock Batch Num]", "'" & .ReturnCode & "'", 1)
            gstrReport.strReportName = Trim$(Left(lstReports, Len(lstReports) - Len(pstrReportType) - 1)) & " for " & .ReturnCode
        End If
        
        If InStr(1, pobjTextBoxThisQuery, "[Account Type]") > 0 Then
            .LabelStr = ""
            .FormCaption = "Account Type Selection"
            .LabelCaption = ""
            .CodeField = "ListCode"
            .DescField = "Description"
            .AddStar = True
            .SQL = "SELECT ListsMaster.ListName, ListDetailsMaster.ListCode, ListDetailsMaster.Description " & _
                "FROM ListsMaster INNER JOIN ListDetailsMaster ON ListsMaster.ListNum = ListDetailsMaster.ListNum " & _
                "WHERE (((ListsMaster.ListName)='Account Type')) " & _
                "ORDER BY ListDetailsMaster.SequenceNum;"
            .Show vbModal
            If .Cancelled = True Then
                FillReplaceParams = True
                Exit Function
            End If
            pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[Account Type]", "'" & .ReturnCode & "'", 1)
            If .ReturnCode <> "*" Then
                gstrReport.strReportName = Trim$(Left(lstReports, Len(lstReports) - Len(pstrReportType) - 1)) & " =  '" & .ReturnCode & "'"
            End If
        End If
        
        If InStr(1, pobjTextBoxThisQuery, "[User ID]") > 0 Then
            .LabelStr = ""
            .FormCaption = "User Selection"
            .LabelCaption = ""
            .CodeField = "UserID"
            .DescField = "UserDesc"
            .AddStar = True
            .SQL = "SELECT UserID, UserName, UserID & ' (' & UserName & ')' as UserDesc From Users ORDER BY UserID;"
            .Show vbModal
            If .Cancelled = True Then
                FillReplaceParams = True
                Exit Function
            End If
            pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[User ID]", "'" & .ReturnCode & "'", 1)
            gstrReport.strReportType = rpTypeDetails
            If .ReturnCode <> "*" Then
                gstrReport.strReportName = Trim$(Left(lstReports, Len(lstReports) - Len(pstrReportType) - 1)) & " =  '" & .ReturnCode & "'"
            End If
        End If
        
        If InStr(1, pobjTextBoxThisQuery, "[Order Status]") > 0 Then
            .LabelStr = ""
            .FormCaption = "Order Status Selection"
            .LabelCaption = ""
            .CodeField = "ListCode"
            .DescField = "Description"
            .AddStar = True
            .SQL = "SELECT ListDetailsMaster.ListCode, ListDetailsMaster.Description FROM ListDetailsMaster INNER JOIN ListsMaster ON ListDetailsMaster.ListNum = ListsMaster.ListNum Where (((ListsMaster.ListName) = 'Order Status')) ORDER BY ListDetailsMaster.SequenceNum;"
            .Show vbModal
            If .Cancelled = True Then
                FillReplaceParams = True
                Exit Function
            End If
            pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[Order Status]", "'" & .ReturnCode & "'", 1)
            gstrReport.strReportType = rpTypeDetails
            If .ReturnCode <> "*" Then
                gstrReport.strReportName = Trim$(Left(lstReports, Len(lstReports) - Len(pstrReportType) - 1)) & " =  '" & .ReturnCode & "'"
            End If
        End If
    End With
    
    If InStr(1, pobjTextBoxThisQuery, "[VAT Rate]") > 0 Then
        pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[VAT Rate]", gstrVATRate, 1)
    End If
    
End Function
Sub AssignReportNames()
Dim lstrImproved As String

    If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
         lstrImproved = "IMPROVED"
    End If

    lstrCreditCardClaim = ColLeveller(Me, lcon2ndColPos, "Credit Card Claims") & gconstrReportTypeTextPrint
    lstrChequeRefundAdviceNotes = ColLeveller(Me, lcon2ndColPos, "Print Refund Cheque AdviceNotes") & gconstrReportTypeTextPrint
    lstrRefundCheques = ColLeveller(Me, lcon2ndColPos, "Refund Cheques") & gconstrReportTypeTextPrint
    
    lstrReportSalesSummary = ColLeveller(Me, lcon2ndColPos, "Sales Summary") & gconstrReportTypeSpreadsheet
    lstrAdviceNoteSpecific = ColLeveller(Me, lcon2ndColPos, "Advice Note Specific") & gconstrReportTypeTextPrint
    lstrAdviceNotes = ColLeveller(Me, lcon2ndColPos, "Advice Notes (A)waiting") & gconstrReportTypeTextPrint
    lstrAdviceNotesRange = ColLeveller(Me, lcon2ndColPos, "Advice Notes (A)waiting Order Range") & gconstrReportTypeTextPrint
    lstrAdviceNotesPrinted = ColLeveller(Me, lcon2ndColPos, "Advice Notes (P)rinted") & gconstrReportTypeTextPrint
    lstrBatchPickings = ColLeveller(Me, lcon2ndColPos, "Batch Pickings List (P)rinted") & gconstrReportTypeTextPrint
    lstrThermalLabels = ColLeveller(Me, lcon2ndColPos, "Thermal Labels") & gconstrReportTypeTextPrint
    
    lstrReportStockAllocatedSummary = ColLeveller(Me, lcon2ndColPos, "Stock (A)llocated or (P)rinted but not Packed.") & gconstrReportTypeSpreadsheet
    lstrReportStockSummary = ColLeveller(Me, lcon2ndColPos, "Stock Out the building, not (D)ownloaded Summary") & gconstrReportTypeSpreadsheet
    
    lstrPrintObjAdviceNotes = ColLeveller(Me, lcon2ndColPos, "Advice Notes (A)waiting   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjRefundAdviceNotes = ColLeveller(Me, lcon2ndColPos, "Refund Cheque Advice Note   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjPFManifest = ColLeveller(Me, lcon2ndColPos, "Parcel Force Manifest   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjRefundCheques = ColLeveller(Me, lcon2ndColPos, "Refund Cheques   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjBatchPick = ColLeveller(Me, lcon2ndColPos, "Batch Pickings List (P)rinted   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjCreditCardClaims = ColLeveller(Me, lcon2ndColPos, "Credit Card Claims   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjCreditCardRefunds = ColLeveller(Me, lcon2ndColPos, "Credit Card Refunds   ") & gconstrReportTypeQuality
    
    lstrPrintObjAdviceNoteSpecific = ColLeveller(Me, lcon2ndColPos, "Advice Note Specific   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjAdviceNotesRange = ColLeveller(Me, lcon2ndColPos, "Advice Notes (A)waiting Order Range   " & lstrImproved) & gconstrReportTypeQuality
    lstrPrintObjAdviceNotesPrinted = ColLeveller(Me, lcon2ndColPos, "Advice Notes (P)rinted   " & lstrImproved) & gconstrReportTypeQuality
    
    lstrPrintObjAdviceNoteStatusSpecific = ColLeveller(Me, lcon2ndColPos, "Advice Note Specific Order Status Export") & gconstrReportTypeQuality
    
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

Sub OrderNumChk(pobjTextOrderNum As Object, pobjLabel As Object)
Dim lstrOrderNum As String
Dim lstrCustName As String

    lstrOrderNum = UCase$(Trim$(pobjTextOrderNum))
    
    If Left(lstrOrderNum, 1) = "M" Then
        lstrOrderNum = Right$(pobjTextOrderNum, Len(lstrOrderNum) - 1)
    End If
    
    If Val(lstrOrderNum) <> 0 Then
        lstrCustName = GetAdviceCustName(CLng(lstrOrderNum))
        If lstrCustName <> "" Then
            pobjLabel = lstrCustName
        Else
            MsgBox "Order not found!", , gconstrTitlPrefix & "Order Number Check"
        End If
    Else
        MsgBox "You must enter a order number", , gconstrTitlPrefix & "Order Number Check"
    End If
    
End Sub

Public Sub cmdBack_Click()

    gstrButtonRoute = gconstrMainMenu
    Set gstrCurrentLoadedForm = frmMainReps
    
    Unload Me
    frmMainReps.Show
    
End Sub

Private Sub cmdEndDate_Click()

    lblEndDate = CheckCalendar(vbKeyInsert, lblEndDate)

End Sub

Private Sub cmdEndOrderNum_Click()

    OrderNumChk txtEndOrderNum, lblCustomerEnd

End Sub

Private Sub cmdExternal_Click()

    MsgBox "The reports normally accessed from External Reporting can now be" & vbCrLf & _
        "accessed from the Reporting program." & vbCrLf & vbCrLf & _
        "Reports are based on data which is live, whereas all data used" & vbCrLf & _
        "External reporting is only updated once a day." & vbCrLf & vbCrLf, vbInformation, gconstrTitlPrefix & "New Features!"

    Busy True, Me
    
    DoEvents
    
    gdatCentralDatabase.Close
    gdatLocalDatabase.Close
    Set gdatLocalDatabase = Nothing
    Set gdatCentralDatabase = Nothing
    
    Busy False, Me
    
    UpdateLoader
    Unload Me
    
    If InStr(UCase(Command$), "/TEST") > 0 Then
        Shell FindProgram("MSACCESS") & " " & gstrStatic.strReportsTestingDBFile, vbNormalFocus
    Else
        Shell FindProgram("MSACCESS") & " " & gstrStatic.strReportsDBFile, vbNormalFocus
    
    End If
        
    End
    
End Sub

Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub
Private Sub cmdPrinterSetup_Click()

    With cdgFile
        .DialogTitle = "Page Setup"
        .Flags = cdlPDPrintSetup ' Or cdlPDNoWarning Or cdlPDReturnIC Or cdlPDReturnDC
        .ShowPrinter
    End With
                
End Sub

Private Sub cmdSelect_Click()
Dim lstrFileName As String
Dim llngOrderNum As Long
Dim lintRetVal As Variant
Dim llngFirstChequeNumber As Long
Dim lbooFoundRecs As Boolean
Dim lbooLocked As Boolean
Dim llngSeqNum As Long
Dim lintInUse As Integer

Const lconWarning = "WARNING: This feature has been replaced and upgraded!" & vbCrLf & vbCrLf & _
    "The data this feature requires is no longer avilable.  " & vbCrLf & vbCrLf & _
    "Please use the 'Quality Printing' version and/or consult Joe Bloggs."
    
    Select Case lstReports.Text
    Case ""
        MsgBox "No report selected", , gconstrTitlPrefix & "Selection"
        Exit Sub
    Case lstrReportStockSummary
        If Not DateOK("END") Then Exit Sub
        cdgFile.FileName = "StockSummary.txt"
        cdgFile.DialogTitle = "Choose export location"
        cdgFile.CancelError = True
        On Error Resume Next
        cdgFile.ShowSave

        If Err.Number <> cdlCancel Then
            DoEvents
            Busy True, Me
            RepStockOrderSummary cdgFile.FileName, 0, CDate(lblEndDate), "BC"
            Busy False, Me
            lintRetVal = MsgBox("Would you like to flag these items as being Dowloaded? ", vbYesNo, gconstrTitlPrefix & "Status Modification")
            If lintRetVal = vbYes Then
                Busy True, Me
                UpdateOrderStatus "D", CDate(lblEndDate), "C"
                UpdateOrderStatus "E", CDate(lblEndDate), "B"
                Busy False, Me
            End If
        Else
            MsgBox "You must select a file name!", , gconstrTitlPrefix & "File Selection"
        End If
    Case lstrAdviceNotes
        If Not DateOK("END") Then Exit Sub
        
        glngItemsWouldLikeToPrint = Val(InputBox("Please enter how many Advice notes you'd like to print! e.g. 20", , 20))
        
        If glngItemsWouldLikeToPrint = 0 Then
            MsgBox "You can't print Zero Advice notes!", , gconstrTitlPrefix & "Mandatory Field"
            Exit Sub
        End If
        
        GetPrinterInfo "LPT2", 66, frmChildPrinter
        Busy True, Me
        lstrFileName = GetTempDir & "I" & Format(Now(), "MMDDSSN")
        PrintAdviceNotesGeneral 0, CDate(lblEndDate), "A", lstrFileName & ".tmp"
        BatchFile lstrFileName
        lintRetVal = MsgBox("Did your Advice Notes Print correctly? " & vbCrLf & _
            "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
        If lintRetVal = vbYes Then
            Busy True, Me
            GetAdviceForPForce 'Create PForce Records
            UpdateOrderStatus "P", CDate(lblEndDate), "A"
            UpdateOrderLineParcelNumber

            Busy False, Me
        End If
        
    Case lstrChequeRefundAdviceNotes
        MsgBox lconWarning, vbInformation, gconstrTitlPrefix & "Upgrade Warning!"
        Exit Sub
        
        If Not DateOK("END") Then Exit Sub
        GetPrinterInfo "LPT1", 66, frmChildPrinter
        Busy True, Me
        lstrFileName = GetTempDir & "R" & Format(Now(), "MMDDSSN")
        PrintChequeRefundAdviceNotes lstrFileName & ".tmp", CDate(lblEndDate)
        BatchFile lstrFileName
        
    Case lstrAdviceNotesPrinted
        If Not DateOK("END") Then Exit Sub
        GetPrinterInfo "LPT2", 66, frmChildPrinter
        Busy True, Me
        lstrFileName = GetTempDir & "I" & Format(Now(), "MMDDSSN")
        PrintAdviceNotesGeneral 0, CDate(lblEndDate), "P", lstrFileName & ".tmp"
        BatchFile lstrFileName
        Busy False, Me

    Case lstrAdviceNoteSpecific
        On Error Resume Next
        llngOrderNum = CLng(InputBox("Please Enter the Order Number for the Advice Note you would like to Print!", "Specific Advice Note Print"))
        On Error GoTo 0
        If llngOrderNum = 0 Then
            MsgBox "This does not appear to be an Order Number!", , gconstrTitlPrefix & "Mandatory Field"
            Exit Sub
        End If
        GetPrinterInfo "LPT2", 66, frmChildPrinter
        Busy True, Me
        lstrFileName = GetTempDir & "I" & Format(Now(), "MMDDSSN")
        PrintAdviceNotesGeneral 0, 0, "S", lstrFileName & ".tmp", llngOrderNum
        BatchFile lstrFileName
        Busy False, Me

    Case lstrReportStockAllocatedSummary
        If Not DateOK("END") Then Exit Sub
        cdgFile.FileName = "StockAwaitSummary.txt"
        cdgFile.DialogTitle = "Choose export location"
        cdgFile.CancelError = True
        On Error Resume Next
        cdgFile.ShowSave
        If Err.Number <> cdlCancel Then
            DoEvents
            Busy True, Me
            RepStockOrderSummary cdgFile.FileName, 0, CDate(lblEndDate), "AP"
            Busy False, Me
        Else
            MsgBox "You must select a file name!", , gconstrTitlPrefix & "File Selection"
        End If
    Case lstrBatchPickings
        Exit Sub
        lstrFileName = GetTempDir & "B" & Format(Now(), "MMDDSSN")
        GetPrinterInfo "LPT2", 66, frmChildPrinter
        Busy True, Me
        lbooFoundRecs = RepBatchPickings(lstrFileName & ".TMP")
        If lbooFoundRecs = True Then
            BatchFile lstrFileName
            Busy False, Me
            lintRetVal = MsgBox("Did your Batch Pickings List print correctly? " & vbCrLf & _
                "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
            If lintRetVal = vbYes Then
                UpdateBatchPrintedFlag
            End If
        End If
        Busy False, Me
    Case lstrCreditCardClaim
        MsgBox lconWarning, vbInformation, gconstrTitlPrefix & "Upgrade Warning!"
        Exit Sub
        
        lstrFileName = GetTempDir & "CC" & Format(Now(), "MMDDSSN")
        GetPrinterInfo "LPT2", 77, frmChildPrinter
        Busy True, Me
        PrintCreditCardClaim lstrFileName & ".TMP"
        BatchFile lstrFileName
        
        lintRetVal = MsgBox("Did CCC Report print OK?", vbYesNo)
        If lintRetVal = vbYes Then
            Busy True, Me
            UpdateAdviceBankRepDate
            Busy False, Me
        End If
        
        Busy False, Me
        
    Case lstrRefundCheques
        MsgBox lconWarning, vbInformation, gconstrTitlPrefix & "Upgrade Warning!"
        Exit Sub
        
        lstrFileName = GetTempDir & "Q" & Format(Now(), "MMDDSSN")
        GetPrinterInfo "LPT2", 77, frmChildPrinter 'Lines not used!
        llngFirstChequeNumber = Val(InputBox("Please enter the first Cheque Number on the first Cheque printed.", "Refund Cheques"))
        If Val(llngFirstChequeNumber) = 0 Then
            MsgBox "You must enter a valid Cheque number!", , gconstrTitlPrefix & "Mandatory Field"
            Exit Sub
        End If
        Busy True, Me
        PrintCheques lstrFileName & ".TMP", llngFirstChequeNumber
        BatchFile lstrFileName
        
        lintRetVal = MsgBox("Did cheques print OK?", vbYesNo)
        If lintRetVal = vbYes Then
            Busy True, Me
            UpdateCheqNumAndPrinted
            Busy False, Me
        End If
        
    Case lstrThermalLabels
        UnloadLastForm
        gstrButtonRoute = gconstrThermalPrintRun
        Set gstrCurrentLoadedForm = frmPForce
        frmPForce.Route = gconstrThermalPrintRun
        frmPForce.CallingForm = frmReports
        mdiMain.DrawButtonSet gstrButtonRoute
        frmPForce.Show

    Case lstrAdviceNotesRange
        If lblCustomerStart <> gconstrReportTypeCustomerNameBlank And lblCustomerEnd <> gconstrReportTypeCustomerNameBlank Then
        
            GetPrinterInfo "LPT2", 66, frmChildPrinter
            Busy True, Me
            lstrFileName = GetTempDir & "I" & Format(Now(), "MMDDSSN")
            PrintAdviceNotesGeneral 0, 0, "R", lstrFileName & ".tmp", CLng(txtStartOrderNum), CLng(txtEndOrderNum)
            BatchFile lstrFileName
            lintRetVal = MsgBox("Did your Advice Notes Print correctly? " & vbCrLf & _
                "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
            If lintRetVal = vbYes Then
                Busy True, Me
                GetAdviceForPForce gconstrAdviceReportTypeRange, CLng(txtStartOrderNum), CLng(txtEndOrderNum)
                UpdateOrderStatus "P", 0, "R", CLng(txtStartOrderNum), CLng(txtEndOrderNum)
                UpdateOrderLineParcelNumber
                Busy False, Me
            End If
                        
        Else
            MsgBox "You must first enter an orders number (in Start and End boxes) " & vbCrLf & _
                "then click button(s) beside them to check that the order exists.", , gconstrTitlPrefix & "Mandatory Field"
            Exit Sub
        End If
    'TESTING ZONE!!!!!
    Case lstrPrintObjAdviceNotes, lstrPrintObjRefundAdviceNotes, lstrPrintObjPFManifest, _
        lstrPrintObjRefundCheques, lstrPrintObjBatchPick, lstrPrintObjCreditCardClaims, _
        lstrPrintObjAdviceNoteSpecific, lstrPrintObjAdviceNotesRange, lstrPrintObjAdviceNotesPrinted, _
        lstrPrintObjRefundAdviceNotes, lstrPrintObjAdviceNoteStatusSpecific, lstrPrintObjCreditCardRefunds
        
        Select Case lstReports.Text
        Case lstrPrintObjAdviceNotes ' A awaiting
            If Not DateOK("END") Then Exit Sub
            
            glngItemsWouldLikeToPrint = Val(InputBox("Please enter how many Advice notes you'd like to print! e.g. 20", , 20))
            
            If glngItemsWouldLikeToPrint = 0 Then
                MsgBox "You can't print Zero Advice notes!", , gconstrTitlPrefix & "Mandatory Field"
                Exit Sub
            End If
            
            ChooseLayout ltAdviceNote, Me
            Busy True, Me
            PrintObjAdviceNotesGeneral 0, CDate(lblEndDate), "A"
            Busy False, Me
            ShowPlotReport
            
            lintRetVal = MsgBox("Did your Advice Notes Print correctly? " & vbCrLf & _
                "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
            If lintRetVal = vbYes Then
                Busy True, Me
                GetAdviceForPForce 'Create PForce Records
                UpdateOrderStatus "P", CDate(lblEndDate), "A"
                UpdateOrderLineParcelNumber
    
                Busy False, Me
            End If
            
        Case lstrPrintObjAdviceNoteStatusSpecific
            Dim lstrOrderStatus As String
            If Not DateOK("END") Then Exit Sub
            
            frmChildOptions.List = "Order Status"
            frmChildOptions.Code = ""
            frmChildOptions.Show vbModal
            lstrOrderStatus = frmChildOptions.Code
    
            gstrReport.booDontDeleteDelim = True
            ChooseLayout ltAdviceNote, Me
            PrintObjAdviceNotesGeneral 0, CDate(lblEndDate), "O", , , lstrOrderStatus
            ShowPlotReport
            
            'export file
            cdgFile.FileName = "AdviceExport.mmos"
            cdgFile.DialogTitle = "Choose export location"
            cdgFile.CancelError = True
            On Error Resume Next
            cdgFile.ShowSave
            If Err.Number <> cdlCancel Then
                Encrypt cdgFile.FileName, gconEncryptDataFile, gstrReport.strDelimDetailsFile
            Else
                MsgBox "You must select a file name!", , gconstrTitlPrefix & "File Selection"
                Exit Sub
            End If
            Kill gstrReport.strDelimDetailsFile
            gstrReport.booDontDeleteDelim = False
            
            lintRetVal = MsgBox("Did your Advice Notes Print correctly? " & vbCrLf & _
                "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
            If lintRetVal = vbYes Then
                Busy True, Me
                UpdateOrderStatus "P", CDate(lblEndDate), "F", , , lstrOrderStatus

                UpdateOrderLineParcelNumber
    
                Busy False, Me
            End If
            
        Case lstrPrintObjAdviceNoteSpecific
            On Error Resume Next
            llngOrderNum = CLng(InputBox("Please Enter the Order Number for the Advice Note you would like to Print!", "Specific Advice Note Print"))
            On Error GoTo 0
            If llngOrderNum = 0 Then
                MsgBox "This does not appear to be an Order Number!", , gconstrTitlPrefix & "Mandatory Field"
                Exit Sub
            End If
            ChooseLayout ltAdviceNote, Me
            Busy True, Me
            PrintObjAdviceNotesGeneral 0, 0, "S", llngOrderNum
            Busy False, Me
            ShowPlotReport

        Case lstrPrintObjAdviceNotesRange
            If lblCustomerStart <> gconstrReportTypeCustomerNameBlank And lblCustomerEnd <> gconstrReportTypeCustomerNameBlank Then
                ChooseLayout ltAdviceNote, Me
                Busy True, Me
                PrintObjAdviceNotesGeneral 0, 0, "R", CLng(txtStartOrderNum), CLng(txtEndOrderNum)
                Busy False, Me
                ShowPlotReport
                
                lintRetVal = MsgBox("Did your Advice Notes Print correctly? " & vbCrLf & _
                    "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
                If lintRetVal = vbYes Then
                    Busy True, Me
                    GetAdviceForPForce gconstrAdviceReportTypeRange, CLng(txtStartOrderNum), CLng(txtEndOrderNum)
                    UpdateOrderStatus "P", 0, "R", CLng(txtStartOrderNum), CLng(txtEndOrderNum)
                    UpdateOrderLineParcelNumber
                    Busy False, Me
                End If
                            
            Else
                MsgBox "You must first enter an orders number (in Start and End boxes) " & vbCrLf & _
                    "then click button(s) beside them to check that the order exists.", , gconstrTitlPrefix & "Mandatory Field"
                Exit Sub
            End If

        Case lstrPrintObjAdviceNotesPrinted
            If Not DateOK("END") Then Exit Sub
            ChooseLayout ltAdviceNote, Me
            Busy True, Me
            PrintObjAdviceNotesGeneral 0, CDate(lblEndDate), "P"
            Busy False, Me
            ShowPlotReport

        Case lstrPrintObjRefundAdviceNotes
            If Not DateOK("END") Then Exit Sub
            ChooseLayout ltAdviceNote, Me
            Busy True, Me
            PrintObjChequeRefundAdviceNotes CDate(lblEndDate)
            Busy False, Me
            ShowPlotReport

        Case lstrPrintObjPFManifest
            ChooseLayout ltParcelForceManifest, Me
            PrintObjPForceManifestGeneral
            Busy False, Me
            ShowPlotReport

        Case lstrPrintObjRefundCheques
            llngFirstChequeNumber = Val(InputBox("Please enter the first Cheque Number on the first Cheque printed.", "Refund Cheques"))
            If Val(llngFirstChequeNumber) = 0 Then
                MsgBox "You must enter a valid Cheque number!", , gconstrTitlPrefix & "Mandatory Field"
                Exit Sub
            End If

            ChooseLayout ltRefundCheques, Me
            Busy True, Me
            PrintObjCheques llngFirstChequeNumber
            Busy False, Me
            ShowPlotReport
            
            lintRetVal = MsgBox("Did cheques print OK?", vbYesNo)
            If lintRetVal = vbYes Then
                Busy True, Me
                UpdateCheqNumAndPrinted
                Busy False, Me
            End If
            
        Case lstrPrintObjBatchPick
            ChooseLayout ltBatchPickings, Me
            lbooFoundRecs = PrintObjBatchPickings
            If lbooFoundRecs = True Then
                Busy False, Me
                ShowPlotReport
                lintRetVal = MsgBox("Did your Batch Pickings List print correctly? " & vbCrLf & _
                    "If so, and you would like flag them as being printed, click YES!", vbYesNo, gconstrTitlPrefix & "Status Modification")
                If lintRetVal = vbYes Then
                    UpdateBatchPrintedFlag
                End If
            End If
            Busy False, Me

        Case lstrPrintObjCreditCardClaims
            ChooseLayout ltCreditCardClaims, Me
            PrintObjCreditCardClaim
            Busy False, Me
            ShowPlotReport
            lintRetVal = MsgBox("Did CCC Report print OK?", vbYesNo)
            If lintRetVal = vbYes Then
                Busy True, Me
                UpdateAdviceBankRepDate
                Busy False, Me
            End If

        Case lstrPrintObjCreditCardRefunds
            ChooseLayout ltCreditCardClaims, Me
            PrintObjCreditCardClaim True
            Busy False, Me
            ShowPlotReport
            lintRetVal = MsgBox("Did CCC Refunds Report print OK?", vbYesNo)
            If lintRetVal = vbYes Then
                Busy True, Me
                UpdateAdviceBankRepDate
                Busy False, Me
            End If
        End Select
    
    Case Else
            txtSQL = ""
            txtThisQuery = ""
            GetCustomRep txtSQL, StripColLevelPadding(lstReports), _
                llngSeqNum, lstrSysDB, lintInUse, lbooLocked, lintOrientation
                        
            If txtSQL = "" Then
                MsgBox "Please choose another!", , gconstrTitlPrefix & "Custom Reporting"
                Exit Sub
            End If
            
            If Not DateOK("S&E") Then Exit Sub
            
            Dim lstrReportType As String
            
            If Right$(lstReports, Len(gconstrReportTypeLabel)) = gconstrReportTypeLabel Then
                lstrReportType = gconstrReportTypeLabel
            ElseIf Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then
                lstrReportType = gconstrReportTypeGrouped
            Else
                lstrReportType = gconstrReportTypeCustom
            End If
            
            ReplaceParams txtSQL, txtThisQuery, lblStartDate, lblEndDate
            
            If FillReplaceParams(txtThisQuery, lstrReportType) = True Then
                Exit Sub
            End If
            
            With gstrReport
                .strReportType = rpTypeDetails
                .strReportName = lstReports
                .booShowOptions = True
            End With
            
            If Right$(lstReports, Len(gconstrReportTypeLabel)) = gconstrReportTypeLabel Then
                AnalyseFields txtThisQuery, lstrSysDB, frmReports
                
                frmChildLabelOptions.PrintingObject = Me
                frmChildLabelOptions.Show vbModal
                If frmChildLabelOptions.Cancelled = True Then Exit Sub
                With gstrReport
                    .strReportType = rpTypeLabels
                    .booHideZoom = True
                    .booShowOptions = False
                End With

            ElseIf Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then
                gstrReport.strReportType = rpTypeGroupings
                'split querys via the ~ character
                txtGroupings = Right$(txtThisQuery, Len(txtThisQuery) - InStr(1, txtThisQuery, "~"))
                txtThisQuery = Left$(txtThisQuery, InStr(1, txtThisQuery, "~") - 1)
                If ReadGroupsBlockIntoArray(txtThisQuery, txtGroupings) = False Then
                    Exit Sub
                End If
                    
                Dim lstrPreMasterSQLCheck As String
                lstrPreMasterSQLCheck = ReplaceStr(txtGroupings, "{", "", 1)
                lstrPreMasterSQLCheck = ReplaceStr(lstrPreMasterSQLCheck, "}", "", 1)
                
                AnalyseFields lstrPreMasterSQLCheck, lstrSysDB, frmReports
            Else
                AnalyseFields txtThisQuery, lstrSysDB, frmReports
            End If
            
            If optOutput(0).Value = True Then
                'Print Preview
                gintScaleFactor = 1
                gbooTotalLineRequired = False
                
                Printer.PaperSize = vbPRPSA4
                
                Font.Name = "Arial"
                Font.Size = 11 / gintScaleFactor
                        
                With gstrReport
                
                    .booShowPageSetup = True
                    .booOptEnableBars = True
                    .booOptEnableLineSpace = True
                    .booOptEnableMargins = True

                    .strStartRangeDate = lblStartDate
                    .strEndRangeDate = lblEndDate

                    .booBarsOn = False
                    .intSpacing = rpSpacing.rpSpacingSingle
                    .lngMargins = rpMargins.rpMarginNarrow
                    .sngFontSize = rpFontFactor.rpFontFactorNormal
                    .strDelimDetailsFile = GetTempDir & "D" & Format(Now(), "MMDDSSN") & ".tmp"
                
                    Dim lintArrInc As Integer
                    Dim lstrBlckHeadSufx As String
                    Dim lintFileNum As Integer
                    If Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then
                        For lintArrInc = 1 To UBound(lstrFieldNames)
                            lstrBlckHeadSufx = lstrBlckHeadSufx & Chr(160) & vbTab
                        Next lintArrInc
                        
                        For lintArrInc = 0 To UBound(mstrMasterGroupSQL)
                                mstrMasterGroupSQL(lintArrInc).strBlockHeader = _
                                    mstrMasterGroupSQL(lintArrInc).strBlockHeader & lstrBlckHeadSufx
                        Next lintArrInc
                        
                        For lintArrInc = 0 To UBound(mstrMasterGroupSQL)
                            lintFileNum = FreeFile
                            Open gstrReport.strDelimDetailsFile For Append As lintFileNum
                            Print #lintFileNum, mstrMasterGroupSQL(lintArrInc).strBlockHeader
                            gstrReport.lngTotalDetailLines = gstrReport.lngTotalDetailLines + 1
                            Close #lintFileNum
                            DoEvents
                            AnalyseSQL mstrMasterGroupSQL(lintArrInc).strSQL, lstrSysDB, frmReports
                        Next lintArrInc
                                    
                    Else
                        AnalyseSQL txtThisQuery, lstrSysDB, frmReports
                    End If
                    
                    If Right$(lstReports, Len(gconstrReportTypeLabel)) = gconstrReportTypeLabel Then
                        If .lngTotalDetailLines <= (gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown) Then
                            .intPagesInReport = 1
                        Else
                            .intPagesInReport = .lngTotalDetailLines / (gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown)
                        End If
                        gstrReport.intDetailLinesOnAPage = 66 ' Used for compatibility only
                    End If
                    
                    Printer.NewPage
                    Printer.EndDoc
                    
                    Printer.Orientation = lintOrientation

                    frmPrintPreview.Show vbModal
                    Set frmPrintPreview = Nothing
                    Kill .strDelimDetailsFile
                End With
                ClearReportBuffer
                ClearReportingDataType
                ClearReportingLayoutType
                ReDim gstrBoxArray(0)
                gintCurrentReportPageNum = 0

            ElseIf optOutput(1).Value = True Then
                'Spreadsheet
                gstrReport.strDelimDetailsFile = StripColLevelPadding(lstReports) & ".txt"
                cdgFile.FileName = gstrReport.strDelimDetailsFile
                cdgFile.DialogTitle = "Choose export location"
                cdgFile.CancelError = True
                On Error Resume Next
                cdgFile.ShowSave

                If Err.Number <> cdlCancel Then
                    If Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then
                        For lintArrInc = 1 To UBound(lstrFieldNames)
                            lstrBlckHeadSufx = lstrBlckHeadSufx & Chr(160) & vbTab
                        Next lintArrInc
                        
                        For lintArrInc = 0 To UBound(mstrMasterGroupSQL)
                                mstrMasterGroupSQL(lintArrInc).strBlockHeader = _
                                    mstrMasterGroupSQL(lintArrInc).strBlockHeader & lstrBlckHeadSufx
                        Next lintArrInc
                        
                        For lintArrInc = 0 To UBound(mstrMasterGroupSQL)
                            lintFileNum = FreeFile
                            Open gstrReport.strDelimDetailsFile For Append As lintFileNum
                            Print #lintFileNum, mstrMasterGroupSQL(lintArrInc).strBlockHeader
                            gstrReport.lngTotalDetailLines = gstrReport.lngTotalDetailLines + 1
                            Close #lintFileNum
                            DoEvents
                            AnalyseSQL mstrMasterGroupSQL(lintArrInc).strSQL, lstrSysDB, frmReports ', mstrMasterGroupSQL(lintArrInc).strBlockHeader
                        Next lintArrInc
                                    
                    Else
                        AnalyseSQL txtThisQuery, lstrSysDB, frmReports ', True
                    End If

                    MsgBox "A Tabbed Delimited File has been created!" & vbCrLf & vbCrLf & "You may now open the file in your spreadsheet package!", , gconstrTitlPrefix & "Spreadsheet file"
                Else
                    MsgBox "You must select a file name!", , gconstrTitlPrefix & "Spreadsheet file"
                End If
            End If
        
            Select Case StripColLevelPadding(lstReports)
            Case "Debit Card Processing Report By User", "Debit Cards Refunds"
                lintRetVal = MsgBox("Did Card Bank Report print OK?", vbYesNo)
                If lintRetVal = vbYes Then
                    Busy True, Me
                    UpdateAdviceBankRepDate
                    Busy False, Me
                End If
        
            End Select
            
    End Select

    Dim lstrReportStr As String
    If Right$(lstReports, Len(gconstrReportTypeLabel)) = gconstrReportTypeLabel Then
        lstrReportStr = "Label"
    ElseIf Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then
        lstrReportStr = "Grouped"
    ElseIf Right$(lstReports, Len(gconstrReportTypeCustom)) = gconstrReportTypeCustom Then
        lstrReportStr = "Custom"
    ElseIf Right$(lstReports, Len(gconstrReportTypeQuality)) = gconstrReportTypeQuality Then
        lstrReportStr = "Quality"
    Else
        lstrReportStr = "Other"
    End If
                
    LogUsage "Report", lstrReportStr, StripColLevelPadding(lstReports)

End Sub

Private Sub cmdStartDate_Click()

    lblStartDate = CheckCalendar(vbKeyInsert, lblStartDate)

End Sub

Private Sub cmdStartOrderNum_Click()

    OrderNumChk txtStartOrderNum, lblCustomerStart

End Sub


Private Sub Command1_Click()

    frmChildLabelOptions.PrintingObject = Me
    
    frmChildLabelOptions.Show vbModal
    If frmChildLabelOptions.Cancelled = True Then Exit Sub

    txtThisQuery = "SELECT trim(trim(DeliverySalutation) & ' ' & trim(DeliveryInitials) & ' ' &  trim(DeliverySurname)) as DelName, DeliveryAdd1, DeliveryAdd2, DeliveryAdd3, DeliveryAdd4, DeliveryAdd5, DeliveryPostcode FROM AdviceNotes WHERE (((DeliverySurname)<>'') AND ((DeliveryAdd1)<>'') AND ((DeliveryAdd2)<>''));"
    
    AnalyseFields txtThisQuery, "CENTRAL", frmReports

    'Print Preview
    gintScaleFactor = 1
    gbooTotalLineRequired = False
    
    Printer.PaperSize = vbPRPSA4
    
    Font.Name = "Arial"
    Font.Size = 11 / gintScaleFactor
            
    With gstrReport
        .booShowPageSetup = True
        .booShowOptions = True
        .booOptEnableBars = True
        .booOptEnableLineSpace = True
        .booOptEnableMargins = True
    
        .strReportType = rpTypeLabels
        .strReportName = lstReports
        .strStartRangeDate = lblStartDate
        .strEndRangeDate = lblEndDate
        .booBarsOn = True
        .intSpacing = rpSpacing.rpSpacingSingle
        .lngMargins = rpMargins.rpMarginNarrow
        .sngFontSize = rpFontFactor.rpFontFactorNormal
        .strDelimDetailsFile = GetTempDir & "D" & Format(Now(), "MMDDSSN") & ".tmp"
    
        AnalyseSQL txtThisQuery, "CENTRAL", frmReports
        
        If .lngTotalDetailLines <= (gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown) Then
            .intPagesInReport = 1
        Else
            .intPagesInReport = .lngTotalDetailLines / (gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown)
        End If
        gstrReport.intDetailLinesOnAPage = 66 ' Used for compatibility only

        frmPrintPreview.Show vbModal
        Set frmPrintPreview = Nothing
        Kill .strDelimDetailsFile
    End With

    ClearReportBuffer
    ClearReportingDataType
    ClearReportingLayoutType
    ReDim gstrBoxArray(0)
    gintCurrentReportPageNum = 0

End Sub
Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    lblEndDate = Format$(Date, "dd/mmm/yyyy")
    
    lblCustomerStart = gconstrReportTypeCustomerNameBlank
    lblCustomerEnd = gconstrReportTypeCustomerNameBlank
    
    If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        cmdExternal.Visible = True
    End If
    
    Select Case gstrGenSysInfo.lngUserLevel
    Case 30 'sales
        FillListBox "Sales"
        TabStrip1.Tabs.Item(4).Selected = True
    Case 40 'Accounts
        FillListBox "Finance"
        TabStrip1.Tabs.Item(2).Selected = True
    Case 50 'General Mangers
        FillListBox "Packing"
        TabStrip1.Tabs.Item(3).Selected = True
    Case Else '99 'IS
        FillListBox "Customer Services"
        TabStrip1.Tabs.Item(1).Selected = True
    End Select
    If DebugVersion Then
        lblStartDate = "01/Jan/1998"
    End If
    
    ShowBanner Me
    
    SetupHelpFileReqs
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngFormHalfWidth As Long

    On Error Resume Next
    llngFormHalfWidth = Me.Width / 2
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdPrinterSetup
        .Top = cmdBack.Top
        .Left = (cmdBack.Left - .Width) - 120
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With

    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
        
    With fraOutput
        .Left = cmdBack.Left
    End With
    
    With cmdSelect
        .Left = cmdBack.Left
    End With
    
    With cmdExternal
        .Left = cmdBack.Left
        .Top = fraOutput.Top + fraOutput.Height + 120
    End With
    
    With TabStrip1
        .Width = cmdBack.Left - 240
        .Height = (cmdBack.Top - TabStrip1.Top) - 425
        lstReports.Left = .ClientLeft + 60
        lstReports.Top = .ClientTop + 60
        lstReports.Height = .ClientHeight - 120
        lstReports.Width = .ClientWidth - 120
    End With
    
    With Label3
        .Top = TabStrip1.Top + TabStrip1.Height
        .Left = 0
        .Width = Me.Width
    End With
    
End Sub

Private Sub lstReports_Click()
Dim lstrParam As String
    
    cmdStartOrderNum.Enabled = False
    cmdEndOrderNum.Enabled = False
    txtStartOrderNum.Enabled = False
    txtEndOrderNum.Enabled = False
    
    If Right$(lstReports, Len(gconstrReportTypeQuality)) = gconstrReportTypeQuality Then: lstrParam = gconstrReportTypeQuality
    If Right$(lstReports, Len(gconstrReportTypeCustom)) = gconstrReportTypeCustom Then: lstrParam = gconstrReportTypeCustom
    If Right$(lstReports, Len(gconstrReportTypeGrouped)) = gconstrReportTypeGrouped Then: lstrParam = gconstrReportTypeGrouped
    If Right$(lstReports, Len(gconstrReportTypeTextPrint)) = gconstrReportTypeTextPrint Then: lstrParam = gconstrReportTypeTextPrint
    If Right$(lstReports, Len(gconstrReportTypeSpreadsheet)) = gconstrReportTypeSpreadsheet Then: lstrParam = gconstrReportTypeSpreadsheet
    If Right$(lstReports, Len(gconstrReportTypeLabel)) = gconstrReportTypeLabel Then: lstrParam = gconstrReportTypeLabel
    
    
    Select Case lstReports.Text
    
    ' Both Start & End Dates
    Case lstrReportSalesSummary
        SetEnabledOutput lstrParam
        cmdEndDate.Enabled = True
        cmdStartDate.Enabled = True
        ShowStatus 16
        
    'Just End Date
    Case lstrReportStockSummary, _
        lstrAdviceNotes, lstrAdviceNotesPrinted, _
        lstrReportStockAllocatedSummary, lstrPrintObjAdviceNotes, lstrPrintObjAdviceNotesPrinted, _
         lstrChequeRefundAdviceNotes, lstrPrintObjRefundAdviceNotes
         SetEnabledOutput lstrParam
        cmdEndDate.Enabled = True
        cmdStartDate.Enabled = False
        ShowStatus 17
        
    'No Dates
    Case lstrAdviceNoteSpecific, lstrBatchPickings, lstrCreditCardClaim, _
         lstrRefundCheques, lstrPrintObjAdviceNoteSpecific, lstrPrintObjBatchPick, _
         lstrThermalLabels, lstrPrintObjCreditCardClaims
         SetEnabledOutput lstrParam
        cmdEndDate.Enabled = False
        cmdStartDate.Enabled = False
        ShowStatus 0
        
    'Rang Dates
    Case lstrAdviceNotesRange, lstrPrintObjAdviceNotesRange
        SetEnabledOutput lstrParam
        txtStartOrderNum.Enabled = True
        txtEndOrderNum.Enabled = True
        cmdStartOrderNum.Enabled = True
        cmdEndOrderNum.Enabled = True
        cmdEndDate.Enabled = False
        cmdStartDate.Enabled = False
        ShowStatus 0
        
    ' No selection, No Dates
    Case ""
        SetEnabledOutput lstrParam
        cmdEndDate.Enabled = False
        cmdStartDate.Enabled = False
        ShowStatus 0
        
    ' Custom reports, Both Start & End Dates
    Case Else
        SetEnabledOutput lstrParam
        cmdEndDate.Enabled = True
        cmdStartDate.Enabled = True
        ShowStatus 109
    End Select
    
End Sub

Private Sub TabStrip1_Click()

    FillListBox TabStrip1.SelectedItem
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub txtEndOrderNum_GotFocus()

    SetSelected Me
    lblCustomerEnd = gconstrReportTypeCustomerNameBlank
    
End Sub

Private Sub txtStartOrderNum_GotFocus()

    SetSelected Me
    lblCustomerStart = gconstrReportTypeCustomerNameBlank
    
End Sub


Sub FillListBox(pstrTabItem As String)

    lstReports.Clear
    'Assign Hard Coded reports
    AssignReportNames
    
    'Add Custom / Label reports from table

    FillCustomReps lstReports, Me, lcon2ndColPos, False, pstrTabItem
    
    Select Case pstrTabItem
    Case "Discrepancies"
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            lstReports.AddItem lstrChequeRefundAdviceNotes
            lstReports.AddItem lstrRefundCheques
        End If
        lstReports.AddItem lstrPrintObjCreditCardRefunds
        lstReports.AddItem lstrPrintObjRefundAdviceNotes
        lstReports.AddItem lstrPrintObjRefundCheques

    Case "Customer Services"
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            lstReports.AddItem lstrCreditCardClaim
        End If
        lstReports.AddItem lstrPrintObjCreditCardClaims

    Case "Finance"
    Case "Packing"
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            lstReports.AddItem lstrAdviceNoteSpecific
            lstReports.AddItem lstrAdviceNotes
            lstReports.AddItem lstrAdviceNotesRange 
            lstReports.AddItem lstrAdviceNotesPrinted
            lstReports.AddItem lstrBatchPickings
            lstReports.AddItem lstrThermalLabels
        End If
        lstReports.AddItem lstrPrintObjAdviceNoteStatusSpecific 
        lstReports.AddItem lstrPrintObjAdviceNoteSpecific 
        lstReports.AddItem lstrPrintObjAdviceNotes
        lstReports.AddItem lstrPrintObjAdviceNotesRange 
        lstReports.AddItem lstrPrintObjAdviceNotesPrinted 
        lstReports.AddItem lstrPrintObjBatchPick
        lstReports.AddItem lstrPrintObjPFManifest
    
    Case "Stock"
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            lstReports.AddItem lstrReportStockAllocatedSummary
            lstReports.AddItem lstrReportStockSummary
        End If
    Case "Sales"
       
    Case "Other"
       
    End Select

End Sub

Sub SetEnabledOutput(pstrParam As String)

    optOutput(0).Enabled = False
    optOutput(1).Enabled = False
    optOutput(2).Enabled = False
    optOutput(3).Enabled = False
    
    Select Case pstrParam
    Case gconstrReportTypeTextPrint
       
    Case gconstrReportTypeSpreadsheet
        optOutput(1).Enabled = True

    Case gconstrReportTypeQuality, gconstrReportTypeLabel
        optOutput(0).Enabled = True
        If optOutput(0).Value = False Then
            optOutput(0).Value = True
        End If

    Case gconstrReportTypeCustom, gconstrReportTypeGrouped
        optOutput(0).Enabled = True
        optOutput(1).Enabled = True
        If optOutput(0).Value = False And optOutput(1).Value = False Then
            optOutput(0).Value = True
        End If

    End Select
        
End Sub

Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/GenReps.xml>WhatsScreen"

    ctlBanner1.WhatsThisHelpID = IDH_GENREPS_MAIN
    ctlBanner1.WhatIsID = IDH_GENREPS_MAIN

    cmdStartDate.WhatsThisHelpID = IDH_GENREPS_STARTDATE
    cmdEndDate.WhatsThisHelpID = IDH_GENREPS_ENDDATE
    txtStartOrderNum.WhatsThisHelpID = IDH_GENREPS_RANGEORDSTART
    cmdStartOrderNum.WhatsThisHelpID = IDH_GENREPS_RANGEORDCMDSTART
    txtEndOrderNum.WhatsThisHelpID = IDH_GENREPS_RANDEORDEND
    cmdEndOrderNum.WhatsThisHelpID = IDH_GENREPS_RANGEORDCMDEND
    TabStrip1.WhatsThisHelpID = IDH_GENREPS_REPSTABS
    lstReports.WhatsThisHelpID = IDH_GENREPS_REPLIST
    cmdSelect.WhatsThisHelpID = IDH_GENREPS_CMDSELECT
    fraOutput.WhatsThisHelpID = IDH_GENREPS_OUTPUT
    optOutput(0).WhatsThisHelpID = IDH_GENREPS_OUTPUT
    optOutput(1).WhatsThisHelpID = IDH_GENREPS_OUTPUT
    optOutput(2).WhatsThisHelpID = IDH_GENREPS_OUTPUT
    optOutput(3).WhatsThisHelpID = IDH_GENREPS_OUTPUT
    cmdPrinterSetup.WhatsThisHelpID = IDH_GENREPS_PRINTSET
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    
End Sub
