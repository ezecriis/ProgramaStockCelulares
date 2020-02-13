VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_agregar_celular 
   Caption         =   "Agregar celular"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Data Data_agregar 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\EXO SMART PRO Q2\Desktop\Profesorado\Mio\ProjecFinalVB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "stock"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txt_cod 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmb_agregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Producto"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc DB_Agregar 
      Height          =   375
      Left            =   600
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
      Connect         =   $"frm_agregar_prod.frx":0000
      OLEDBString     =   $"frm_agregar_prod.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Agregar"
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
      Caption         =   "Marca:"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frm_agregar_celular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_agregar_Click()

cod = txt_cod

If txt_cod = Empty Or txt_prod = Empty Or txt_pre = Empty Or txt_stock = Empty Then

variable = MsgBox("Debe ingresar todos los datos", , "Error")
txt_cod.SetFocus

Else

DB_Agregar.Recordset.FindFirst ("Codigo =" + cod)

If DB_Agregar.Recordset.NoMatch Then

DB_Agregar.Recordset.AddNew

DB_Agregar.Recordset!codigo = txt_cod
DB_Agregar.Recordset!producto = txt_prod
DB_Agregar.Recordset!precio = txt_pre
DB_Agregar.Recordset!stock = txt_stock

DB_Agregar.Recordset.Update

variable = MsgBox("Producto guardado exitosamente", , "Resultado del alta")

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus

Else

variable = MsgBox("Producto ya existente", , "Resultado del alta")

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus
End If
End If
DB_Agregar.Recordset.MoveFirst

End Sub

Private Sub cmb_volver_Click()
Me.Hide

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

End Sub

txt_cod.SetFocus
DB_Agregar.Refresh
End Sub

