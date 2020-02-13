VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_listado 
   Caption         =   "Lista de productos"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data_lista 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\EXO SMART PRO Q2\Desktop\Profesorado\Mio\ProjecFinalVB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   585
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "stock"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Bindings        =   "frm_listado.frx":0000
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
   End
   Begin MSAdodcLib.Adodc DB_Lista 
      Height          =   375
      Left            =   840
      Top             =   4440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Connect         =   $"frm_listado.frx":0019
      OLEDBString     =   $"frm_listado.frx":00AF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Lista"
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
End
Attribute VB_Name = "frm_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_volver_Click()
Me.Hide
End Sub
Private Sub Form_Activate()
DB_Lista.Refresh
End Sub

Private Sub Form_Load()

flex.TextMatrix(0, 0) = "Codigo"
flex.TextMatrix(0, 1) = "Producto"
flex.TextMatrix(0, 2) = "Precio"
flex.TextMatrix(0, 3) = "Stock"

End Sub

