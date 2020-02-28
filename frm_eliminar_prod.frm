VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_eliminar_equipo 
   Caption         =   "Eliminar equipo"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form3"
   ScaleHeight     =   6960
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmb_limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
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
      DataSource      =   "DB_Eliminar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "DB_Eliminar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Modelos"
      DataSource      =   "DB_Eliminar"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txt_imei 
      DataField       =   "Imei"
      DataSource      =   "DB_Eliminar"
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
      OLEDBString     =   $"frm_eliminar_prod.frx":0094
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Modelos"
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
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "IMEI:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frm_eliminar_equipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_eliminar_Click()

producto = txt_imei

If producto = Empty Then

respuesta = MsgBox("Debe buscar un producto", , "Error")

Else

respuesta = MsgBox("¿Eliminar producto?", vbYesNo, "Confirmar eliminacion")

If respuesta = 6 Then

DB_Eliminar.Recordset.Delete

variable = MsgBox("Producto eliminado exitosamente", , "Resultado de la eliminacion")

txt_imei = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_imei.SetFocus

Else

variable = MsgBox("Eliminacion cancelada", , "Resultado de eliminacion")

txt_imei = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_imei.SetFocus

End If
End If

End Sub

Private Sub cmb_limpiar_Click()

txt_imei = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

txt_imei.SetFocus

End Sub

Private Sub cmb_volver_Click()
Me.Hide

txt_imei = ""
txt_prod = ""
txt_pre = ""
txt_stock = ""

End Sub

Private Sub cmd_buscar_Click()
On Error GoTo salida
DB_Eliminar.Recordset.MovePrevious
If DB_Eliminar.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = InputBox("Ingrese el imei a buscar:", "Sistema de registro")
DB_Eliminar.Recordset.Find "IMEI='" & Trim(busqueda) & "'"
If DB_Eliminar.Recordset.EOF Then
MsgBox "El imei ingresado no se ha encontrado", vbCritical, "Sistema de Registro"
Exit Sub
End If
txt_prod.Text = DB_Eliminar.Recordset.Fields(0).Value
txt_imei.Text = DB_Eliminar.Recordset.Fields(1).Value
txt_pre.Text = DB_Eliminar.Recordset.Fields(2).Value
txt_stock.Text = DB_Eliminar.Recordset.Fields(3).Value

Exit Sub
salida:
MsgBox "llenar campo", vbCritical, "Sistema de registro"
End Sub

Private Sub Form_Activate()
txt_imei.SetFocus
DB_Eliminar.Refresh
End Sub
