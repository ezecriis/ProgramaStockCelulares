VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_modificar_prod 
   Caption         =   "Modificar producto"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmb_Limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmb_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmb_modificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "DB_Modificar"
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "DB_Modificar"
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1770
      Width           =   975
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Producto"
      DataSource      =   "DB_Modificar"
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1170
      Width           =   1455
   End
   Begin VB.TextBox txt_cod 
      DataField       =   "Codigo"
      DataSource      =   "DB_Modificar"
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin MSAdodcLib.Adodc DB_Modificar 
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
      Connect         =   $"frm_modificar_prod.frx":0000
      OLEDBString     =   $"frm_modificar_prod.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Modificar"
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
      Caption         =   "Stock"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Producto:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frm_modificar_prod"
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

DB_Modificar.Recordset.FindFirst ("Codigo =" + cod)

If DB_Modificar.Recordset.NoMatch Then
variable = MsgBox("Producto no encontrado", , "Resultado de modificacion")
txt_cod = ""
txt_cod.SetFocus

Else
txt_prod = DB_Modificar.Recordset!producto
txt_pre = DB_Modificar.Recordset!precio
txt_stock = DB_Modificar.Recordset!stock

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

Private Sub cmb_modificar_Click()

respuesta = MsgBox("¿Modificar producto?", vbYesNo, "Confirmar modificacion")

If respuesta = 6 Then

DB_Modificar.Recordset.Edit

DB_Modificar.Recordset!producto = txt_prod
DB_Modificar.Recordset!precio = txt_pre
DB_Modificar.Recordset!stock = txt_stock

variable = MsgBox("Producto modificado exitosamente", , "Resultado de modificacion")

DB_Modificar.Recordset.Update

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus

Else

variable = MsgBox("Modificacion cancelada", , "Resultado de modificacion")

txt_cod = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_cod.SetFocus

End If

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
DB_Modificar.Refresh
DB_Modificar.Recordset.Edit
DB_Modificar.Recordset.Update
End Sub
