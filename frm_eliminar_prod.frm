VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_eliminar_prod 
   Caption         =   "Eliminar producto"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form3"
   ScaleHeight     =   7170
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmb_limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Data Data_eliminar 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\EXO SMART PRO Q2\Desktop\Profesorado\Mio\ProjecFinalVB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "stock"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmb_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmb_eliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Producto"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txt_cod 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin MSAdodcLib.Adodc DB_Eliminar 
      Height          =   375
      Left            =   480
      Top             =   4560
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
      Connect         =   $"frm_eliminar_prod.frx":0000
      OLEDBString     =   $"frm_eliminar_prod.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Eliminar"
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
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Producto:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frm_eliminar_prod"
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

DB_Eliminar.Recordset.FindFirst ("Codigo = " + cod)

If Adodc1.Recordset.NoMatch Then

variable = MsgBox("Producto no encontrado", , "Resultado de la eliminacion")

txt_cod = ""
txt_cod.SetFocus

Else

txt_prod = DB_Eliminar.Recordset!producto
txt_pre = DB_Eliminar.Recordset!precio
txt_stock = DB_Eliminar.Recordset!stock

End If
End If

End Sub

Private Sub cmb_eliminar_Click()

producto = txt_prod

If producto = Empty Then

respuesta = MsgBox("Debe buscar un producto", , "Error")

Else

respuesta = MsgBox("¿Eliminar producto?", vbYesNo, "Confirmar eliminacion")

If respuesta = 6 Then

DB_Eliminar.Recordset.Delete

variable = MsgBox("Producto eliminado exitosamente", , "Resultado de la eliminacion")

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus

Else

variable = MsgBox("Eliminacion cancelada", , "Resultado de eliminacion")

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus

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
DB_Eliminar.Refresh

End Sub
