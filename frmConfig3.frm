VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig3 
   Caption         =   "System Configuration Wizard"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmConfig3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "B&rowse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox cboBarangay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmConfig3.frx":08CA
      Left            =   4800
      List            =   "frmConfig3.frx":092B
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1080
      Left            =   5040
      Picture         =   "frmConfig3.frx":0A94
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Official Logo (Optional):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   120
      Picture         =   "frmConfig3.frx":112F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registry of Barangay Inhabitants version 1.2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   4140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfig3.frx":67B1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   1305
   End
End
Attribute VB_Name = "frmConfig3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
        
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo cam_trap
    AddFilter CommonDialog1, "Image Files", "*.gif;*.jpg;*.jpeg;*.png"
    CommonDialog1.ShowOpen
    If Not CommonDialog1.FileName = "" Then
save_again:
        FileCopy CommonDialog1.FileName, App.Path & "\Resident\brgy_logo.jpg"
        Image2 = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
    End If
    Exit Sub
cam_trap:
        If Err.Number = 76 Then
            MkDir App.Path & "\Resident"
            GoTo save_again
        End If
End Sub

Private Sub AddFilter(ByVal dlg As CommonDialog, ByVal filter_title As String, ByVal filter_value As String)
Dim txt As String

    txt = dlg.Filter
    If Len(txt) > 0 Then txt = txt & "|"
    txt = txt & filter_title & " (" & filter_value & ")|" & _
        filter_value
    dlg.Filter = txt
End Sub

Private Sub Command1_Click()
    Load frmConfig2
    frmConfig2.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
    If cboBarangay.ListIndex = -1 Then
        MsgBox "Please select the name of your barangay.", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        bn = cboBarangay.List(cboBarangay.ListIndex)
        Load frmConfig4
        frmConfig4.Show
        Me.Hide
    End If
End Sub

Private Sub Command3_Click()
    Dim choice As Integer
    choice = MsgBox("The system has not been configured successfully." & vbNewLine & "Are you sure you want to terminate this?", vbYesNo + vbQuestion, "Registry of Barangay Inhabitants")
    If choice = vbYes Then
        End
    End If
End Sub
