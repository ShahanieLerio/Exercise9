VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form repCustomerInfo 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18105
   LinkTopic       =   "Form1"
   ScaleHeight     =   12525
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Go Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtFilter 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   10455
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   17175
      _cx             =   30295
      _cy             =   18441
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   13321
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "repCustomerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Sub PrintReport()

        '<EhHeader>
    On Error GoTo PrintReport_Err

        '</EhHeader>

    Dim CRApp As New CRAXDDRT.Application
    Dim Report As CRAXDDRT.Report

    Set Report = CRApp.OpenReport(ReportPath & "repCustomersInfo.rpt")
    Report.Database.Tables(1).Location = ServerPath & "\DB\NewDB.mdb"
    'Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
     CRViewer.ReportSource = Report

    
    Dim strString As String
   

    MousePointer = vbHourglass

    
    ' Use US-style date string in Crystal syntax
    
    strString = "{tblCustomer.Name} = '" & txtFilter.Text & "'"
    
                    
  
    Report.RecordSelectionFormula = strString

   
    cmdPrint.Enabled = False
    cmdPrint.Caption = "Working"

    CRViewer.ViewReport

    Do While CRViewer.IsBusy
        DoEvents
    Loop

    CRViewer.Zoom 100
    Me.WindowState = vbMaximized


    cmdPrint.Enabled = True
    cmdPrint.Caption = "Go Print"
    MousePointer = vbNormal
  
  
    '<EhFooter>
        Exit Sub

PrintReport_Err:
        ErrReport Err.Description, _
            "Please call brayan immediately 0915-891-8530 LendingClientV2.rep_LoansMaturityChecker.PrintReport", _
            Erl

        Resume Next

        '/EhFooter>

End Sub

Private Sub cmdPrint_Click()
    PrintReport
End Sub
