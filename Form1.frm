VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "H:\College\5th Semester\SAD\Practical 7\prac7.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LC"
      Top             =   9480
      Width           =   5295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   30
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GENERATE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   29
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   28
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   27
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   26
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   25
      Top             =   7320
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Branch"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3600
      TabIndex        =   23
      Text            =   "Select Branch"
      Top             =   6480
      Width           =   3615
   End
   Begin VB.TextBox Text9 
      DataField       =   "Leaving Reason"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      DataField       =   "Leaving Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      DataField       =   "Previous College"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   4800
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      DataField       =   "Admission Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8520
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "Birth Place"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "DOB"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Country"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Text            =   "Select Country"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Caste"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   7920
      TabIndex        =   8
      Text            =   "Select Caste"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "Mother's Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leaving Reason :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leaving Date :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Admission Date :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Previous College :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Birth Place :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOB :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Country :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caste :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mother's Name :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student ID :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEAVING CERTIFICATE GENERATOR"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Data1.Recordset.Edit
End Sub

Private Sub Command3_Click()
Data1.Recordset.Delete
End Sub

Private Sub Command4_Click()
Data1.Recordset.Update
End Sub

Private Sub Command5_Click()
DataReport1.Show
End Sub

Private Sub Command6_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Combo1.AddItem ("OPEN")
Combo1.AddItem ("OBC")
Combo1.AddItem ("VJNT")
Combo1.AddItem ("SC")
Combo1.AddItem ("ST")
Combo1.AddItem ("SBC")
Combo2.AddItem ("India")
Combo3.AddItem ("Computer Science and Engineering")
Combo3.AddItem ("Information Technology")
Combo3.AddItem ("Mechanical Engineering")
Combo3.AddItem ("Civil Engineering")
Combo3.AddItem ("Electrical Engineering")
Combo3.AddItem ("Electronics and Telecommunication Engineering")
Combo3.AddItem ("Instrumentation Engineering")
End Sub
