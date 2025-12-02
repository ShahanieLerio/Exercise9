VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21270
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   21270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCustomerInfo 
      Caption         =   "Customer's Info"
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
      Left            =   11040
      TabIndex        =   30
      Top             =   5400
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   600
      TabIndex        =   28
      Top             =   6000
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   6800
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   1920
      TabIndex        =   27
      Top             =   5400
      Width           =   8895
   End
   Begin VB.TextBox txtContNum 
      Height          =   495
      Left            =   10440
      TabIndex        =   19
      Top             =   2880
      Width           =   5415
   End
   Begin VB.TextBox txtStatus 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   20175
      Begin VB.TextBox txtGender 
         Height          =   495
         Left            =   1680
         TabIndex        =   29
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   18240
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   18240
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   16440
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   16440
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtBusiness 
         Height          =   495
         Left            =   9840
         TabIndex        =   21
         Top             =   1560
         Width           =   5415
      End
      Begin VB.TextBox txtAge 
         Height          =   495
         Left            =   13800
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtBirthdate 
         Height          =   495
         Left            =   9840
         TabIndex        =   15
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135462913
         CurrentDate     =   45974
      End
      Begin VB.TextBox txtNationality 
         Height          =   495
         Left            =   9840
         TabIndex        =   13
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox txtAddress 
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox txtName 
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000003&
         Caption         =   "Business:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000003&
         Caption         =   "Contact No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000003&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   "Birthdate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000003&
         Caption         =   "Nationality:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000003&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txtCode 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "CODE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "CUSTOMER'S DETAILS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()
On Error GoTo btnAdd_Click_Err

        'TxtLog "Entered btnAdd_Click"
        '</EhHeader>
     If btnAdd.Caption = "Add" Then
         btnAdd.Caption = "Save"
         btnClose.Caption = "Cancel"
         txtName.Enabled = True
         txtAddress.Enabled = True
         txtGender.Enabled = True
         txtStatus.Enabled = True
         dtBirthdate.Enabled = True
         txtAge.Enabled = True
         txtContNum.Enabled = True
         txtBusiness.Enabled = True
         txtNationality.Enabled = True
         txtName.SetFocus
         DataGrid1.Enabled = True
        Else

128             If MsgBox( _
                    "Are you sure to you want to add this record?", _
                    vbQuestion + vbYesNo, _
                    "J Lending Corporation") = _
                    vbYes Then
    

130                With rsCustomer2New
                        .AddNew
                        .Fields("Name") = txtName.Text
                        .Fields("Address") = txtAddress.Text
                        .Fields("Gender") = txtGender.Text
                        .Fields("Status") = txtStatus.Text
                        .Fields("Birthdate") = dtBirthdate.Value
                        .Fields("Age") = Val(txtAge.Text)
                        .Fields("ContNumber") = txtContNum.Text
                        .Fields("Business") = txtBusiness.Text
                        .Fields("Nationality") = txtNationality.Text
                        .Update
                   End With

                     
                     MsgBox _
                            "Record Successfully Added", _
                            vbInformation, _
                            "Webplus Lending Corporation"
                    Unload frmCustomer
                    Me.Show
                    txtSearch.Enabled = True
                

            End If
        End If

        '<EhFooter>

        Exit Sub
btnAdd_Click_Err:
        ErrReport Err.Description, "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.btnAdd_Click", Erl

        Resume Next


        '</EhFooter>
End Sub

Private Sub btnClose_Click()
  Unload Me
  Me.Show
End Sub

Private Sub btnDelete_Click()

'<EhHeader>
        On Error GoTo btnDelete_Click_Err

116         If MsgBox( _
                "Are you sure you want to delete this user?", _
                vbQuestion + vbYesNo) = vbYes _
                Then
118             rsCustomer2New.Delete
120             rsCustomer2New.Update
122             MsgBox _
                    "Record successfully deleted", _
                    vbInformation
124             Unload Me
126             Me.Show
            End If
 

        '<EhFooter>
        Exit Sub

btnDelete_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.btnDelete_Click", _
            Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()
  
   
         '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        'TxtLog "Entered btnEdit_Click"
        '</EhHeader>
100     If btnEdit.Caption = "Edit" Then
102         btnEdit.Caption = "Update"
104         txtName.Enabled = True
106         txtAddress.Enabled = True
108         txtGender.Enabled = True
110         txtStatus.Enabled = True
114         dtBirthdate.Enabled = True
116         DataGrid1.Enabled = False
118     ElseIf btnEdit.Caption = "Update" Then
                btnEdit.Caption = "Edit"
120         If rsCustomer2New.State = 1 Then rsCustomer2New.Close
122         rsCustomer2New.Open _
                "Select * from tblCustomer where Name = '" _
                & txtName.Text & "'"

124         If txtName.Text = "" Or _
                txtAddress.Text = "" Or _
                txtGender.Text = "" Or _
                txtStatus.Text = "" Then
126             MsgBox _
                    "All fields are required! ", _
                    vbInformation
            Else

128             If MsgBox( _
                    "Are you sure to update this record?", _
                    vbQuestion + vbYesNo, _
                    "J Lending Corporation") = _
                    vbYes Then

                Call Customer2New
                    
130                 With rsCustomer2New
136                     !Name = _
                            txtName.Text
138                     !Address = _
                            txtAddress.Text
140                     !Gender = _
                            txtGender.Text
142                     !Status = _
                            txtStatus.Text
144                     !Birthdate = _
                            dtBirthdate.Value
                        !Age = _
                            txtAge.Text
                        !ContNumber = _
                            txtContNum.Text
                        !Business = _
                            txtBusiness.Text
                        !Nationality = _
                            txtNationality.Text
148                     .Update
                     
                     MsgBox _
                            "Record Successfully Updated", _
                            vbInformation, _
                            "Webplus Lending Corporation"
                    Unload frmCustomer
                    Me.Show
                    End With

                End If
            End If
        End If

        '<EhFooter>
        'TxtLog "Exited btnEdit_Click"
        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.btnEdit_Click", _
            Erl

        Resume Next


        '</EhFooter>

End Sub

Private Sub Command1_Click()
    Load repCustomerInfo
    repCustomerInfo.Show
End Sub

Private Sub cmdCustomerInfo_Click()

        '<EhHeader>
        On Error GoTo cmdCustomerInfo_Click_Err

        'TxtLog "Entered cmdSOAStatement_Click"
        '</EhHeader>
        If txtSearch.Text = "" Then
            MsgBox "Type Search first"
            txtSearch.Enabled = True
        End If

        txtSearch.SetFocus

        If rsCustomer2New.RecordCount = 0 Then Exit Sub

100     With rsCustomer2New
102         Unload repCustomerInfo

104         'If rsCustomer2New.EOF Then
                'MsgBox "Please search customer first................. "
           'Else
108            ' Call repCustomerInfo
           ' End If

110         repCustomerInfo.Show
        End With

        '<EhFooter>
        'TxtLog "Exited  cmdCustomerInfo_Click"
        Exit Sub

cmdCustomerInfo_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_payment. cmdCustomerInfo_Click"
           

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        On Error GoTo DataGrid1_Click_Err

100     If rsCustomer2New.RecordCount = 0 Then
        Else
102         btnAdd.Enabled = False
104         btnClose.Caption = "Cancel"
106         btnEdit.Enabled = True
108         btnDelete.Enabled = True

110         With rsCustomer2New
118             txtName.Text = IIf(IsNull(!Name), "", !Name)
119             txtAddress.Text = IIf(IsNull(!Address), "", !Address)
120             txtGender.Text = IIf(IsNull(!Gender), "", !Gender)
121             txtStatus.Text = IIf(IsNull(!Status), "", !Status)
122
123         If IsNull(!Birthdate) Then
124            dtBirthdate.Value = Date
125         Else
126             dtBirthdate.Value = !Birthdate
127         End If

128             txtAge.Text = IIf(IsNull(!Age), "0", !Age)
129             txtContNum.Text = IIf(IsNull(!ContNumber), "", !ContNumber)
130             txtBusiness.Text = IIf(IsNull(!Business), "", !Business)
131             txtNationality.Text = IIf(IsNull(!Nationality), "", !Nationality)
            End With

        End If

     
        Exit Sub
     
DataGrid1_Click_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.DataGrid1_Click", _
            Erl

        Resume Next
End Sub


Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

       
100     Call connect
        'Calls recordset for Customer

102     Call Customer2New
116     Me.Show
118     Me.SetFocus
120     Me.WindowState = vbMaximized

122     If rsCustomer2New.State = 1 Then _
            rsCustomer2New.Close
124         rsCustomer2New.Open _
            "Select * from tblCustomer Order By Code desc"
            
128     Set DataGrid1.DataSource = rsCustomer2New
        'Adjusting the width of Fields on Datagird
130     DataGrid1.Width = 19500
132     DataGrid1.Columns(0).Width = 550
134     DataGrid1.Columns(1).Width = 550
136     DataGrid1.Columns(2).Width = 950
138     DataGrid1.Columns(3).Width = 4000
139     DataGrid1.Columns(4).Width = 2000
140     DataGrid1.Columns(5).Width = 1100
        DataGrid1.Columns(6).Width = 550
141     DataGrid1.Columns(7).Width = 950
142     DataGrid1.Columns(8).Width = 4000
143     DataGrid1.Columns(9).Width = 2000


        '<EhFooter>
        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, _
            "Please call brayan immediately 0915-891-8530 LendingClientV2.frm_Customer.Form_Load", _
            Erl

        Resume Next

        '</EhFooter>
End Sub

Private Sub txtSearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        'TxtLog "Entered txtsearch_Change"
        '</EhHeader>
100     If rsCustomer2New.State = 1 Then rsCustomer2New.Close
102     rsCustomer2New.Open _
           "Select * from tblCustomer where Name like '" _
                & txtSearch.Text & _
                "%' or Address like '" & _
                txtSearch.Text & _
                "%' or Gender like '" & _
                txtSearch.Text & _
                "' Order by Code desc"
             Set DataGrid1.DataSource = rsCustomer2New

        '<EhFooter>
        'TxtLog "Exited txtsearch_Change"
        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, _
            "LendingClientV2.frm_Users.txtsearch_Change", _
            Erl

        Resume Next
        
        
        '</EhFooter>
End Sub
