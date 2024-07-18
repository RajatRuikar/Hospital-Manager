VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form updatedoctor 
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   18
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   5640
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8040
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\swara\Desktop\Hospital mgt\db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\swara\Desktop\Hospital mgt\db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "doctor"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "JOIN DATE"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1320
      TabIndex        =   17
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   16080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   3360
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2520
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   2040
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "QUALIFICATION"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SPECIALLIZATION"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE DOCTOR"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   10
      Top             =   0
      Width           =   11055
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   0
      Left            =   0
      Picture         =   "updatedoctor.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   11895
   End
End
Attribute VB_Name = "updatedoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Adodc1.Recordset.Fields("ID") = Text1.Text
Adodc1.Recordset.Fields("dname") = Text2.Text
Adodc1.Recordset.Fields("dadd") = Text3.Text
Adodc1.Recordset.Fields("mob") = Text4.Text
Adodc1.Recordset.Fields("dquali") = Text5.Text
Adodc1.Recordset.Fields("specialization") = Text6.Text
Adodc1.Recordset.Fields("jdate") = Text7.Text
Adodc1.Recordset.Update
MsgBox "Data Updated"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command3_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Command4_Click()
On Error GoTo errmsg
Command1.Visible = True
Adodc1.Refresh
 Adodc1.Recordset.Find "ID=" & Val(Text1.Text)


    

'Text2.Text = Adodc1.Recordset.Fields("pid")
Text2.Text = Adodc1.Recordset.Fields("dname")
Text3.Text = Adodc1.Recordset.Fields("dadd")
Text4.Text = Adodc1.Recordset.Fields("mob")
Text5.Text = Adodc1.Recordset.Fields("dquali")
Text6.Text = Adodc1.Recordset.Fields("specialization")
Text7.Text = Adodc1.Recordset.Fields("jdate")
 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub
