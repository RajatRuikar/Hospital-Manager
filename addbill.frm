VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addbill 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   8880
      TabIndex        =   41
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   8880
      TabIndex        =   40
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   8880
      TabIndex        =   39
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   8880
      TabIndex        =   38
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   975
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
      TabIndex        =   2
      Top             =   7080
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
      Left            =   4320
      TabIndex        =   1
      Top             =   7080
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
      TabIndex        =   0
      Top             =   7080
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7080
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
      RecordSource    =   "bill"
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
      Left            =   7080
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   7080
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7080
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
      RecordSource    =   "bill"
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
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text12 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      TabIndex        =   23
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
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
      TabIndex        =   24
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
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
      TabIndex        =   25
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Charges"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   6360
      TabIndex        =   37
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital charge"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   6360
      TabIndex        =   36
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Equipment charges"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   5880
      TabIndex        =   35
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Medical/surgical"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   6240
      TabIndex        =   34
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   1335
      Index           =   2
      Left            =   9480
      TabIndex        =   16
      Top             =   0
      Width           =   2055
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
      Height          =   735
      Index           =   6
      Left            =   480
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   600
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   720
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
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
      Height          =   615
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   4320
      Width           =   735
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
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill NO"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   16080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW BILL"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   9135
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   0
      Left            =   0
      Picture         =   "addbill.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   1
      Left            =   0
      Picture         =   "addbill.frx":29787
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   3
      Left            =   0
      Picture         =   "addbill.frx":52F0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image imgLogo 
      Height          =   9465
      Index           =   2
      Left            =   0
      Picture         =   "addbill.frx":7C695
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW BILL"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   33
      Top             =   0
      Width           =   9135
   End
   Begin VB.Line Line2 
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
      Height          =   735
      Index           =   13
      Left            =   1560
      TabIndex        =   32
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
      Height          =   615
      Index           =   12
      Left            =   960
      TabIndex        =   31
      Top             =   3480
      Width           =   1095
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
      Height          =   615
      Index           =   11
      Left            =   1080
      TabIndex        =   30
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   720
      TabIndex        =   29
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   600
      TabIndex        =   28
      Top             =   5760
      Width           =   1335
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
      Height          =   735
      Index           =   8
      Left            =   480
      TabIndex        =   27
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   1335
      Index           =   7
      Left            =   8280
      TabIndex        =   26
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "addbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields("ID") = Text1.Text
 Adodc1.Recordset.Fields("pid") = Text2.Text
Adodc1.Recordset.Fields("pname") = Text3.Text
'Adodc1.Recordset.Fields("dadd") = Text3.Text
Adodc1.Recordset.Fields("pmob") = Text4.Text
Adodc1.Recordset.Fields("disease") = Text5.Text
Adodc1.Recordset.Fields("amt") = Text6.Text
Adodc1.Recordset.Fields("date") = Label2(2).Caption
Adodc1.Recordset.Fields("room") = Text13.Text
Adodc1.Recordset.Fields("hospital") = Text14.Text
Adodc1.Recordset.Fields("equipment") = Text15.Text
Adodc1.Recordset.Fields("medical") = Text16.Text
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
End Sub

Private Sub Command2_Click()
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

Private Sub Form_Load()
Label2(2).Caption = Time & vbTab & Date
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
'Text3.Text = Adodc1.Recordset.Fields("padd")
Text4.Text = Adodc2.Recordset.Fields("mob")
Text5.Text = Adodc2.Recordset.Fields("disease")
'Text6.Text = Adodc1.Recordset.Fields("pdate")

 Exit Sub
errmsg:
MsgBox "record not exist"
End Sub

Private Sub Text6_Click()
Dim b, e, f, g, h, i, j, k, l, m, n, o, p
b = 0
k = 0
l = 0
m = 0
n = 0
o = 0
f = 0
g = 0
h = 0
i = 0
j = 0

e = Text13.Text
e = e + 1
f = Text14.Text
f = f + 1
g = Text15.Text
g = g + 1
h = Text16.Text
h = h + 1


Text6.Text = e + f + g + h - 4
End Sub
