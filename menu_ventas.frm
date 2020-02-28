VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form menu_ventas 
   Caption         =   "Ventas"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   6960
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmb_limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmb_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmb_realizar 
      Caption         =   "Realizar"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txt_cant 
      DataSource      =   "DB_Venta"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "DB_Venta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "DB_Venta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Producto"
      DataSource      =   "DB_Venta"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txt_codigo 
      DataField       =   "Codigo"
      DataSource      =   "DB_Venta"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin MSAdodcLib.Adodc DB_Venta 
      Height          =   375
      Left            =   600
      Top             =   4800
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
      Connect         =   $"menu_ventas.frx":0000
      OLEDBString     =   $"menu_ventas.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "productos"
      Caption         =   "DB Ventas"
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
   Begin VB.Label Label5 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Producto:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "menu_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_buscar_Click()

cod = txt_codigo

If txt_codigo = Empty Then

variable = MsgBox("Debe ingresar un codigo", , "Error")
txt_codigo.SetFocus

Else

DB_Venta.Recordset.FindFirst ("Codigo =" + cod)

If DB_Venta.Recordset.NoMatch Then

variable = MsgBox("Producto no encontrado", , "Resultado de la venta")
txt_codigo = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""
txt_codigo.SetFocus

Else

txt_prod = DB_Venta.Recordset!producto
txt_pre = DB_Venta.Recordset!precio
txt_stock = DB_Venta.Recordset!stock

txt_cant.SetFocus

End If
End If

End Sub

Private Sub cmb_limpiar_Click()

txt_codigo = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""
txt_cant = ""
txt_codigo.SetFocus

End Sub

Private Sub cmb_realizar_Click()

cantidad = txt_cant
stock = txt_stock

If cantidad = Empty Then
variable = MsgBox("Debe ingresar una cantidad", , "Error")
txt_cant.SetFocus

Else

x = stock - cantidad

If x < 0 Then
variable = MsgBox("No tiene suficiente stock", , "Error")
txt_cant.SetFocus

Else

total = stock - cantidad

DB_Venta.Recordset.Edit

DB_Venta.Recordset!stock = total

DB_Venta.Recordset.Update

variable = MsgBox("Venta realizada exitosamente", , "Resultado de la venta")

txt_codigo = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""
txt_cant = ""
txt_codigo.SetFocus

End If
End If

End Sub

Private Sub cmb_volver_Click()
Me.Hide

txt_codigo = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

End Sub

Private Sub Form_Activate()
txt_codigo.SetFocus
DB_Venta.Refresh

End Sub
