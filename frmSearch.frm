VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   7200
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbDatabase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbDatabase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblRBI"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10215
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   8295
         Begin VB.Label lblResult 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Result:"
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
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1275
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSearch.frx":08CA
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Left            =   8520
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
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
         Left            =   8520
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find Now"
         Default         =   -1  'True
         DownPicture     =   "frmSearch.frx":08DF
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
         Left            =   8520
         Picture         =   "frmSearch.frx":0ABE
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8295
         Begin VB.TextBox txtKeyword 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            TabIndex        =   1
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cboFilter 
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
            ItemData        =   "frmSearch.frx":0C9D
            Left            =   6000
            List            =   "frmSearch.frx":0CB0
            TabIndex        =   2
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Filter:"
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
            Left            =   4560
            TabIndex        =   9
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Keyword:"
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
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   1305
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3960
      Picture         =   "frmSearch.frx":0CEF
      Top             =   -360
      Width           =   6465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
      Height          =   240
      Left            =   960
      TabIndex        =   12
      Top             =   240
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmSearch.frx":1AAC
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
    Dim sql As String
    If txtKeyword = "" Then
        MsgBox "Please enter keyword.", vbExclamation, "Registry of Barangay Inhabitants"
    ElseIf cboFilter.ListIndex = -1 Then
        MsgBox "Please select filter.", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        If cboFilter.ListIndex = 0 Then
            sql = "SELECT * FROM tblRBI WHERE Surname LIKE '" & txtKeyword & "%'"
        ElseIf cboFilter.ListIndex = 1 Then
            sql = "SELECT * FROM tblRBI WHERE FirstName LIKE '" & txtKeyword & "%'"
        ElseIf cboFilter.ListIndex = 2 Then
            sql = "SELECT * FROM tblRBI WHERE MiddleName LIKE '" & txtKeyword & "%'"
        ElseIf cboFilter.ListIndex = 3 Then
            sql = "SELECT * FROM tblRBI WHERE Street LIKE '" & txtKeyword & "%'"
        ElseIf cboFilter.ListIndex = 4 Then
            sql = "SELECT * FROM tblRBI WHERE Occupation LIKE '" & txtKeyword & "%'"
        End If
        Adodc1.RecordSource = sql
        Adodc1.Refresh
        DataGrid1.Refresh
        lblResult = "Search Result:" & " " & Adodc1.Recordset.RecordCount & " Record(s) found"
    End If
End Sub

Private Sub cmdView_Click()
    If Adodc1.Recordset.EOF Then
        MsgBox "No record to display, Please try again!", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        frmMain.txtResident = Adodc1.Recordset.Fields("Resident")
        frmMain.txtRelationship = Adodc1.Recordset.Fields("Relationship")
        frmMain.txtSurname = Adodc1.Recordset.Fields("Surname")
        frmMain.txtFirstName = Adodc1.Recordset.Fields("FirstName")
        frmMain.txtMiddleName = Adodc1.Recordset.Fields("MiddleName")
        frmMain.txtStreet = Adodc1.Recordset.Fields("Street")
        frmMain.cboBarangay = Adodc1.Recordset.Fields("Barangay")
        frmMain.cboBirthdate = Adodc1.Recordset.Fields("Birthdate")
        frmMain.txtBirthplace = Adodc1.Recordset.Fields("Birthplace")
        frmMain.cboAttainment = Adodc1.Recordset.Fields("Attainment")
        frmMain.txtCitizenship = Adodc1.Recordset.Fields("Citizenship")
        frmMain.txtOccupation = Adodc1.Recordset.Fields("Occupation")
        frmMain.cboBloodType = Adodc1.Recordset.Fields("BloodType")
        frmMain.txtFSurname = Adodc1.Recordset.Fields("FSurname")
        frmMain.txtFFirstName = Adodc1.Recordset.Fields("FFirstName")
        frmMain.txtFMiddleName = Adodc1.Recordset.Fields("FMiddleName")
        frmMain.txtMSurname = Adodc1.Recordset.Fields("MSurname")
        frmMain.txtMFirstName = Adodc1.Recordset.Fields("FirstName")
        frmMain.txtMMiddleName = Adodc1.Recordset.Fields("MMiddleName")
        If Adodc1.Recordset.Fields("Relationship") = "" Then
            frmMain.optYes1 = True
            txtRelationship = ""
        Else
            frmMain.optNo1 = True
            txtRelationship = Adodc1.Recordset.Fields("Relationship")
        End If
        If Adodc1.Recordset.Fields("Gender") = "Male" Then
            frmMain.optMale = True
        Else
            frmMain.optFemale = True
        End If
        If Adodc1.Recordset.Fields("Civil") = "Single" Then
            frmMain.optSingle = True
        ElseIf Adodc1.Recordset.Fields("Civil") = "Married" Then
            frmMain.optMarried = True
        ElseIf Adodc1.Recordset.Fields("Civil") = "Widow/er" Then
            frmMain.optWidow = True
        ElseIf Adodc1.Recordset.Fields("Civil") = "Separated" Then
            frmMain.optSeparated = True
        End If
        If Adodc1.Recordset.Fields("PWDs") = "No" Then
            frmMain.optNo2 = True
        Else
            frmMain.optYes2 = True
        End If
        If Adodc1.Recordset.Fields("Voter") = "No" Then
            frmMain.optNo3 = True
        Else
            frmMain.optYes3 = True
        End If
        If Adodc1.Recordset.Fields("House") = "Concrete" Then
            frmMain.optConcrete = True
        ElseIf Adodc1.Recordset.Fields("House") = "Semi Concrete" Then
            frmMain.optSemiConcrete = True
        ElseIf Adodc1.Recordset.Fields("House") = "Make Shift" Then
            frmMain.optMakeShift = True
        End If
        frmMain.txtAge = DateDiff("yyyy", Adodc1.Recordset.Fields("Birthdate"), Now)
        Me.Hide
        Unload Me
    End If
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub
