VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form eliminar_imei2 
   Caption         =   "Eliminar por IMEI"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   720
      Top             =   4800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Connect         =   $"eliminar_imei2.frx":0000
      OLEDBString     =   $"eliminar_imei2.frx":0094
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Modelos"
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
   Begin VB.CommandButton cmb_eliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox DB_Modelo 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmd_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmb_limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txt_stock 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txt_pre 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txt_cod 
      DataField       =   "Imei"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txt_prod 
      DataField       =   "Modelos"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "IMEI:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "eliminar_imei2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_eliminar_Click()

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

Private Sub cmd_buscar_Click()
On Error GoTo salida
DB_Modelo.Recordset.MovePrevious
If DB_Modelo.Recordset.BOF Then
End If
Dim busqueda As String
busqueda = InputBox("Ingrese el imei a buscar:", "Sistema de registro")
DB_Modelo.Recordset.Find "Imei='" & Trim(busqueda) & "'"
If DB_Modelo.Recordset.EOF Then
MsgBox "El imei no se ha encontrado", vbCritical, "Sistema de Registro"
Exit Sub
End If
txt_prod.Text = DB_Modelo.Recordset.Fields(0).Value
txt_cod.Text = DB_Modelo.Recordset.Fields(1).Value
txt_pre.Text = DB_Modelo.Recordset.Fields(2).Value
txt_stock.Text = DB_Modelo.Recordset.Fields(3).Value

Exit Sub
salida:
MsgBox "llenar campo", vbCritical, "Sistema de registro"
End Sub

Private Sub Form_Activate()
txt_cod.SetFocus
DB_Modelo.Refresh
End Sub

