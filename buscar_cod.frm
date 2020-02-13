VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form buscar_modelo 
   Caption         =   "Busqueda por modelo"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc DB_Modelo 
      Height          =   375
      Left            =   360
      Top             =   4080
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
      Connect         =   $"buscar_cod.frx":0000
      OLEDBString     =   $"buscar_cod.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Modelo ex code"
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
   Begin VB.Data Data_cod 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\EXO SMART PRO Q2\Desktop\Profesorado\Mio\ProjecFinalVB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "stock"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmb_limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Producto"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmb_buscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txt_cod 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Marca:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IMEI:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "buscar_modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_buscar_Click()

cod = txt_cod

If txt_cod = Empty Then

variable = MsgBox("Debe ingresar un codigo", , "Error")
txt_cod.SetFocus

Else

Data_cod.Recordset.FindFirst ("Codigo = " + cod)

If Data_cod.Recordset.NoMatch Then

variable = MsgBox("Producto no encontrado", , "Resultado de busqueda")

txt_cod = ""
txt_cod.SetFocus

Else

txt_prod = Data_cod.Recordset!producto
txt_pre = Data_cod.Recordset!precio
txt_stock = Data_cod.Recordset!stock

End If
End If

End Sub

Private Sub cmb_limpiar_Click()

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus
End Sub

Private Sub cmb_volver_Click()
Me.Hide

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

End Sub

Private Sub Form_Activate()
txt_cod.SetFocus
DB_Modelo.Refresh
End Sub
