VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form updatecovid 
   Caption         =   "Form2"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form2"
   ScaleHeight     =   9465
   ScaleWidth      =   11325
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
      Index           =   1
      Left            =   3600
      TabIndex        =   16
      Top             =   2400
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
      Left            =   7080
      TabIndex        =   8
      Top             =   7080
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
      Left            =   4320
      TabIndex        =   7
      Top             =   7080
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
      Left            =   1560
      TabIndex        =   6
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5880
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8880
      Top             =   6360
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
      RecordSource    =   "covid"
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
      Left            =   8760
      Top             =   6000
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE COVID PATIENT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11895
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   16080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   0
      Left            =   0
      Picture         =   "updatecovid.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "updatecovid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Adodc1.Recordset.Fields("pid") = Text2.Text
Adodc1.Recordset.Fields("pname") = Text3.Text
Adodc1.Recordset.Fields("padd") = Text4.Text
Adodc1.Recordset.Fields("pmob") = Text5.Text
Adodc1.Recordset.Fields("adate") = Text6.Text
Adodc1.Recordset.Update
MsgBox "Data Updated"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command3_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Command4_Click(Index As Integer)
On Error GoTo errmsg
Command1.Visible = True
Adodc1.Refresh
 Adodc1.Recordset.Find "ID=" & Val(Text1.Text)


    

Text2.Text = Adodc1.Recordset.Fields("pid")
Text3.Text = Adodc1.Recordset.Fields("pname")
Text4.Text = Adodc1.Recordset.Fields("padd")
Text5.Text = Adodc1.Recordset.Fields("pmob")
Text6.Text = Adodc1.Recordset.Fields("adate")

 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub
