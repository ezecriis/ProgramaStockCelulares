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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   600
      Top             =   5040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      OLEDBString     =   $"frm_agregar_prod.frx":0094
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txt_marca 
      DataField       =   "Imei"
      DataSource      =   "DB_Agregar"
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmb_agregar 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
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
   Begin VB.TextBox txt_modelo 
      DataField       =   "Modelos"
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
      Connect         =   $"frm_agregar_prod.frx":0128
      OLEDBString     =   $"frm_agregar_prod.frx":01BC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Modelos"
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
Private Sub BtnNuevo_Click()
DB_Agregar.Recordset.AddNew
cmb_agregar.Enabled = True
End Sub

Private Sub cmb_agregar_Click()



If txt_marca.Text <> "" Or txt_modelo.Text <> "" Or txt_pre.Text <> "" Or txt_stock.Text <> "" Then
DB_Agregar.Recordset.Update
Else
mensaje = MsgBox("Completar casillas que esten en  blanco", vbCritical, "Completar")
End If
End Sub

'cod = txt_marca

'If txt_marca = Empty Or txt_modelo = Empty Or txt_pre = Empty Or txt_stock = Empty Then

'variable = MsgBox("Debe ingresar todos los datos", , "Error")
'txt_marca.SetFocus

'Else

'DB_Agregar.Recordset.FindFirst ("Codigo =" + cod)

'If DB_Agregar.Recordset.NoMatch Then

'DB_Agregar.Recordset.AddNew

'DB_Agregar.Recordset!marca = txt_marca
'DB_Agregar.Recordset!modelo = txt_modelo
'DB_Agregar.Recordset!precio = txt_pre
'DB_Agregar.Recordset!stock = txt_stock

'DB_Agregar.Recordset.Update

'variable = MsgBox("Producto guardado exitosamente", , "Resultado del alta")

'txt_marca = ""
'txt_modelo = ""
'txt_pre = ""
'txt_stock = ""

'txt_marca.SetFocus

'Else

'variable = MsgBox("Producto ya existente", , "Resultado del alta")

'txt_marca = ""
'txt_modelo = ""
'txt_pre = ""
'txt_stock = ""

'txt_marca.SetFocus
'End If
'End If
'DB_Agregar.Recordset.MoveFirst

'End Sub

Private Sub cmb_volver_Click()
Me.Hide

txt_marca = ""
txt_modelo = ""
txt_pre = ""
txt_stock = ""

End Sub



'Private Sub Form_Activate()
'txt_marca.SetFocus
'DB_Agregar.Refresh
'End Sub
Private Sub Form_Load()

End Sub
