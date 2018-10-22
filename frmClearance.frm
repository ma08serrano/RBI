VERSION 5.00
Begin VB.Form frmClearance 
   Caption         =   "Clearance"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   Icon            =   "frmClearance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4215
      Begin VB.TextBox txtPurpose 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtRequestor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtResident 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purpose:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requestor Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resident Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   4215
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   2160
      Picture         =   "frmClearance.frx":08CA
      Top             =   0
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLEARANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmClearance.frx":0FE9
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmClearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()
    On Error Resume Next
    Dim sql As String
    If txtResident = "" Or txtRequestor = "" Or txtPurpose = "" Then
        MsgBox "Please fill up all the required fields.", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        sql = "SELECT * FROM tblRBI WHERE Resident=" & txtResident
        Set rs = db.Execute(sql)
        If rs.EOF Then
            MsgBox "No such record found, Please try again!", vbExclamation, "Registry of Barangay Inhabitants"
        Else
            Load drClearance
            Set drClearance.DataSource = rs
            drClearance.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
            drClearance.Sections(2).Controls("lblCaptain").Caption = "HON. " & UCase(GetSetting(App.EXEName, "Config", "Captain"))
            drClearance.Sections(2).Controls("lblSecretary").Caption = UCase(GetSetting(App.EXEName, "Config", "Secretary"))
            drClearance.Sections(2).Controls("lblName").Caption = rs!Surname & ", " & rs!FirstName & " " & Left(rs!MiddleName, 1) & "."
            drClearance.Sections(2).Controls("lblAge").Caption = DateDiff("yyyy", rs!Birthdate, Now)
            drClearance.Sections(2).Controls("lblAddress").Caption = rs!Street & " " & rs!Barangay & ", " & rs!City
            drClearance.Sections(2).Controls("lblRequestor").Caption = txtRequestor
            drClearance.Sections(2).Controls("lblMonth").Caption = Format(Now, "MMMM")
            drClearance.Sections(2).Controls("lblDay").Caption = Format(Now, "DD")
            drClearance.Sections(2).Controls("lblYear").Caption = Format(Now, "yy")
            drClearance.Sections(2).Controls("lblPurpose").Caption = txtPurpose
            Dim res As Integer
            sql = "SELECT * FROM tblClearance ORDER BY ResidentCertNo DESC"
            Set rs = db.Execute(sql)
            If rs.EOF Then
                res = 1
            Else
                res = rs.Fields(0) + 1
            End If
            sql = "INSERT INTO tblClearance VALUES(" & res & ",'" & Format(Now, "MM/DD/YYYY") & "'," & txtResident & ",'" & txtRequestor & "','" & txtPurpose & "')"
            db.Execute sql
            drClearance.Sections(2).Controls("lblCert").Caption = res
            drClearance.Sections(2).Controls("lblNo").Caption = res
            drClearance.Sections(2).Controls("lblDate").Caption = Format(Now, "MM/DD/YYYY")
            If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
                Set drClearance.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
            Else
                drClearance.Sections(1).Controls("Image2").Visible = False
            End If
            drClearance.Show vbModal
            Me.Hide
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub
