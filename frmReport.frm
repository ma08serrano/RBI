VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   4215
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
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox txtAuthor 
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
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
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
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4215
      Begin VB.TextBox txtTitle 
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
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
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
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      Begin VB.ComboBox cboMatrix 
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
         ItemData        =   "frmReport.frx":08CA
         Left            =   1800
         List            =   "frmReport.frx":08E3
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select population:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT"
      Height          =   240
      Left            =   960
      TabIndex        =   12
      Top             =   240
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmReport.frx":0941
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   2160
      Picture         =   "frmReport.frx":0CE9
      Top             =   0
      Width           =   2310
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()
    Dim sql As String
    If cboMatrix.ListIndex = -1 Then
        MsgBox "Please select population.", vbExclamation, "Registry of Barangay Inhabitants"
    End If
    If cboMatrix.ListIndex = 0 Then
        Dim age1, age3, age4, age5, age6, age7, age8, age9, age10, age11, age12, age13, age14, age15, age16, age17, age18 As Integer
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<5"
        Set rs = db.Execute(sql)
        age1 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<10 AND DateDiff('yyyy',Birthdate, Now)>4"
        Set rs = db.Execute(sql)
        age3 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<15 AND DateDiff('yyyy',Birthdate, Now)>9"
        Set rs = db.Execute(sql)
        age4 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<20 AND DateDiff('yyyy',Birthdate, Now)>14"
        Set rs = db.Execute(sql)
        age5 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<25 AND DateDiff('yyyy',Birthdate, Now)>19"
        Set rs = db.Execute(sql)
        age6 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<30 AND DateDiff('yyyy',Birthdate, Now)>24"
        Set rs = db.Execute(sql)
        age7 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<35 AND DateDiff('yyyy',Birthdate, Now)>29"
        Set rs = db.Execute(sql)
        age8 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<40 AND DateDiff('yyyy',Birthdate, Now)>34"
        Set rs = db.Execute(sql)
        age9 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<45 AND DateDiff('yyyy',Birthdate, Now)>39"
        Set rs = db.Execute(sql)
        age10 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<50 AND DateDiff('yyyy',Birthdate, Now)>44"
        Set rs = db.Execute(sql)
        age11 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<55 AND DateDiff('yyyy',Birthdate, Now)>49"
        Set rs = db.Execute(sql)
        age12 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<60 AND DateDiff('yyyy',Birthdate, Now)>54"
        Set rs = db.Execute(sql)
        age13 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<65 AND DateDiff('yyyy',Birthdate, Now)>59"
        Set rs = db.Execute(sql)
        age14 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<70 AND DateDiff('yyyy',Birthdate, Now)>64"
        Set rs = db.Execute(sql)
        age15 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<75 AND DateDiff('yyyy',Birthdate, Now)>69"
        Set rs = db.Execute(sql)
        age16 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)<80 AND DateDiff('yyyy',Birthdate, Now)>74"
        Set rs = db.Execute(sql)
        age17 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE DateDiff('yyyy',Birthdate, Now)>79"
        Set rs = db.Execute(sql)
        age18 = rs.Fields(0)
        Load drAge
        Set drAge.DataSource = rs
        drAge.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drAge.Sections(2).Controls("lblTitle").Caption = txtTitle
        drAge.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drAge.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drAge.Sections(2).Controls("lblMatrix1").Caption = "0 - 4: " & age1
        drAge.Sections(2).Controls("lblMatrix3").Caption = "5 - 9: " & age3
        drAge.Sections(2).Controls("lblMatrix4").Caption = "10 - 14: " & age4
        drAge.Sections(2).Controls("lblMatrix5").Caption = "15 - 19: " & age5
        drAge.Sections(2).Controls("lblMatrix6").Caption = "20 - 24: " & age6
        drAge.Sections(2).Controls("lblMatrix7").Caption = "25 - 29: " & age7
        drAge.Sections(2).Controls("lblMatrix8").Caption = "30 - 34: " & age8
        drAge.Sections(2).Controls("lblMatrix9").Caption = "35 - 39: " & age9
        drAge.Sections(2).Controls("lblMatrix10").Caption = "40 - 44: " & age10
        drAge.Sections(2).Controls("lblMatrix11").Caption = "45 - 49: " & age11
        drAge.Sections(2).Controls("lblMatrix12").Caption = "50 - 54: " & age12
        drAge.Sections(2).Controls("lblMatrix13").Caption = "55 - 59: " & age13
        drAge.Sections(2).Controls("lblMatrix14").Caption = "60 - 64: " & age14
        drAge.Sections(2).Controls("lblMatrix15").Caption = "65 - 69: " & age15
        drAge.Sections(2).Controls("lblMatrix16").Caption = "70 - 74: " & age16
        drAge.Sections(2).Controls("lblMatrix17").Caption = "75 - 79: " & age17
        drAge.Sections(2).Controls("lblMatrix18").Caption = "80+: " & age18
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drAge.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drAge.Sections(1).Controls("Image2").Visible = False
        End If
        drAge.Sections(2).Controls("lblTotal").Caption = "Total Population: " & age1 + age3 + age4 + age5 + age6 + age7 + age8 + age9 + age10 + age11 + age12 + age13 + age14 + age15 + age16 + age17 + age18
        drAge.Show vbModal
    ElseIf cboMatrix.ListIndex = 1 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total1, male As Integer
        total1 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Gender ='Male'"
        Set rs = db.Execute(sql)
        male = rs.Fields(0)
        Load drGender
        Set drGender.DataSource = rs
        drGender.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drGender.Sections(2).Controls("lblTitle").Caption = txtTitle
        drGender.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drGender.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drGender.Sections(2).Controls("lblMale").Caption = "Male: " & male
        drGender.Sections(2).Controls("lblFemale").Caption = "Female: " & (total1 - male)
        drGender.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total1
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drGender.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drGender.Sections(1).Controls("Image2").Visible = False
        End If
        drGender.Show vbModal
    ElseIf cboMatrix.ListIndex = 2 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total2, stat1, stat2, stat3, stat4 As Integer
        total2 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Civil ='Single'"
        Set rs = db.Execute(sql)
        stat1 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Civil ='Married'"
        Set rs = db.Execute(sql)
        stat2 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Civil ='Widow/er'"
        Set rs = db.Execute(sql)
        stat3 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Civil ='Separated'"
        Set rs = db.Execute(sql)
        stat4 = rs.Fields(0)
        Load drCivilStatus
        Set drCivilStatus.DataSource = rs
        drCivilStatus.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drCivilStatus.Sections(2).Controls("lblTitle").Caption = txtTitle
        drCivilStatus.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drCivilStatus.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drCivilStatus.Sections(2).Controls("lblSingle").Caption = "Single: " & stat1
        drCivilStatus.Sections(2).Controls("lblMarried").Caption = "Married: " & stat2
        drCivilStatus.Sections(2).Controls("lblWidow").Caption = "Widow/er: " & stat3
        drCivilStatus.Sections(2).Controls("lblSeparated").Caption = "Separated: " & stat4
        drCivilStatus.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total2
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drCivilStatus.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drCivilStatus.Sections(1).Controls("Image2").Visible = False
        End If
        drCivilStatus.Show vbModal
    ElseIf cboMatrix.ListIndex = 3 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total3, yes As Integer
        total3 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Voter = 'Yes'"
        Set rs = db.Execute(sql)
        yes = rs.Fields(0)
        Load drRegisteredVoter
        Set drRegisteredVoter.DataSource = rs
        drRegisteredVoter.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drRegisteredVoter.Sections(2).Controls("lblTitle").Caption = txtTitle
        drRegisteredVoter.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drRegisteredVoter.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drRegisteredVoter.Sections(2).Controls("lblYes").Caption = "Yes: " & yes
        drRegisteredVoter.Sections(2).Controls("lblNo").Caption = "No: " & (total3 - yes)
        drRegisteredVoter.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total3
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drRegisteredVoter.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drRegisteredVoter.Sections(1).Controls("Image2").Visible = False
        End If
        drRegisteredVoter.Show vbModal
    ElseIf cboMatrix.ListIndex = 4 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total4, yes2 As Integer
        total4 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE PWDs = 'Yes'"
        Set rs = db.Execute(sql)
        yes2 = rs.Fields(0)
        Load drPWDs
        Set drPWDs.DataSource = rs
        drPWDs.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drPWDs.Sections(2).Controls("lblTitle").Caption = txtTitle
        drPWDs.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drPWDs.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drPWDs.Sections(2).Controls("lblYes").Caption = "Yes: " & yes2
        drPWDs.Sections(2).Controls("lblNo").Caption = "No: " & (total4 - yes2)
        drPWDs.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total4
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drPWDs.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drPWDs.Sections(1).Controls("Image2").Visible = False
        End If
        drPWDs.Show vbModal
    ElseIf cboMatrix.ListIndex = 5 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total5, concrete, semiconcrete, makeshift As Integer
        total5 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE House = 'Concrete'"
        Set rs = db.Execute(sql)
        concrete = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE House = 'Semi Concrete'"
        Set rs = db.Execute(sql)
        semiconcrete = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI Where House = 'Make Shift'"
        Set rs = db.Execute(sql)
        makeshift = rs.Fields(0)
        Load drHouse
        Set drHouse.DataSource = rs
        drHouse.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drHouse.Sections(2).Controls("lblTitle").Caption = txtTitle
        drHouse.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drHouse.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drHouse.Sections(2).Controls("lblConcrete").Caption = "Concrete: " & concrete
        drHouse.Sections(2).Controls("lblSemiConcrete").Caption = "Semi Concrete: " & semiconcrete
        drHouse.Sections(2).Controls("lblMakeShift").Caption = "Make Shift: " & makeshift
        drHouse.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total5
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drHouse.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drHouse.Sections(1).Controls("Image2").Visible = False
        End If
        drHouse.Show vbModal
    ElseIf cboMatrix.ListIndex = 6 Then
        sql = "SELECT COUNT(*) FROM tblRBI"
        Set rs = db.Execute(sql)
        Dim total6, elementary, secondary, vocational, degree, masteral, doctoral As Integer
        total6 = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Elementary'"
        Set rs = db.Execute(sql)
        elementary = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Secondary'"
        Set rs = db.Execute(sql)
        secondary = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Vocational'"
        Set rs = db.Execute(sql)
        vocational = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Degree Course'"
        Set rs = db.Execute(sql)
        degree = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Masteral Degree'"
        Set rs = db.Execute(sql)
        masteral = rs.Fields(0)
        sql = "SELECT COUNT(*) FROM tblRBI WHERE Attainment = 'Doctoral Degree'"
        Set rs = db.Execute(sql)
        doctoral = rs.Fields(0)
        Load drAttainment
        Set drAttainment.DataSource = rs
        drAttainment.Sections(1).Controls("lblBrgy").Caption = "BARANGAY " & UCase(GetSetting(App.EXEName, "Config", "Barangay Name"))
        drAttainment.Sections(2).Controls("lblTitle").Caption = txtTitle
        drAttainment.Sections(2).Controls("lblDate").Caption = Format(Now, "MMMM DD, YYYY")
        drAttainment.Sections(2).Controls("lblAuthor").Caption = "Prepared by: " & txtAuthor
        drAttainment.Sections(2).Controls("lblElementary").Caption = "Elementary: " & elementary
        drAttainment.Sections(2).Controls("lblSecondary").Caption = "Secondary: " & secondary
        drAttainment.Sections(2).Controls("lblVocational").Caption = "Vocational: " & vocational
        drAttainment.Sections(2).Controls("lblDegreeCourse").Caption = "Degree Course: " & degree
        drAttainment.Sections(2).Controls("lblMasteralDegree").Caption = "Masteral Degree: " & masteral
        drAttainment.Sections(2).Controls("lblDoctoralDegree").Caption = "Doctoral Degree: " & doctoral
        drAttainment.Sections(2).Controls("lblTotal").Caption = "Total Population: " & total6
        If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
           Set drAttainment.Sections(1).Controls("Image2").Picture = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
        Else
            drAttainment.Sections(1).Controls("Image2").Visible = False
        End If
        drAttainment.Show vbModal
        Me.Hide
        Unload Me
    End If
End Sub

Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub
