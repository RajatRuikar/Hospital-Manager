VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addpathology 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
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
      Left            =   6720
      TabIndex        =   2
      Top             =   8520
      Width           =   1695
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
      TabIndex        =   1
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   8520
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8520
      Top             =   8280
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
      RecordSource    =   "pathology"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8520
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "patient"
      Caption         =   "Adodc2"
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
      Caption         =   "PATIENT ID"
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
      Index           =   7
      Left            =   1080
      TabIndex        =   18
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PATHOLOGY"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST PRISE"
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
      Left            =   1320
      TabIndex        =   16
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST NAME"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   5520
      Width           =   2895
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
      Left            =   1920
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
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
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   2295
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
      Left            =   2160
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
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
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   735
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
      Caption         =   "TEST DATE"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   0
      Left            =   0
      Picture         =   "addpathology.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "addpathology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields("ID") = Text1.Text
 Adodc1.Recordset.Fields("pid") = Text2.Text
Adodc1.Recordset.Fields("pname") = Text3.Text
Adodc1.Recordset.Fields("padd") = Text4.Text
Adodc1.Recordset.Fields("pmob") = Text5.Text
Adodc1.Recordset.Fields("tname") = Text6.Text
Adodc1.Recordset.Fields("tprise") = Text7.Text
Adodc1.Recordset.Fields("tdate") = Text8.Text

Adodc1.Recordset.Update
MsgBox "Data Added"
Adodc1.Recordset.MoveLast
Text1.Text = Adodc1.Recordset.Fields("ID") + 1
Text1.Enabled = False
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Command3_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.Recordset.MoveLast
Text1.Text = Adodc1.Recordset.Fields("ID") + 1
Text1.Enabled = False

End Sub

Private Sub Text3_Click()
On Error GoTo errmsg
Command1.Visible = True
Adodc2.Refresh
 Adodc2.Recordset.Find "ID=" & Val(Text2.Text)


    

'Text1.Text = Adodc1.Recordset.Fields("ID")
Text3.Text = Adodc2.Recordset.Fields("pname")
Text4.Text = Adodc2.Recordset.Fields("padd")
Text5.Text = Adodc2.Recordset.Fields("mob")
'Text5.Text = Adodc2.Recordset.Fields("disease")
'Text6.Text = Adodc1.Recordset.Fields("pdate")

 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub
