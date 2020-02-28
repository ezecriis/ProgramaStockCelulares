VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_Agregar 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtimei 
      DataField       =   "Imei"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"frm_Agregar.frx":0000
      OLEDBString     =   $"frm_Agregar.frx":0094
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
   Begin VB.CommandButton BtnVolver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnAgr 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnNvo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtsto 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtpre 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtmod 
      DataField       =   "Modelos"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtmarca 
      DataField       =   "Marca"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Imei:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Marca:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Agregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgr_Click()
If txtmarca.Text <> "" Or txtmod.Text <> "" Or txtpre.Text <> "" Or txtsto.Text <> "" Or txtimei.Text <> "" Then
Adodc1.Recordset.Update
mensa = MsgBox("El equipo se cargo con éxito", vbCritical, "Carga equipo")
Else
mensaje = MsgBox("Completar las casillas que esten en blanco", vbCritical, "Completar")
End If
End Sub

Private Sub BtnNvo_Click()
Adodc1.Recordset.AddNew
BtnAgr.Enabled = True
End Sub

Private Sub BtnVolver_Click()
frm_Agregar.Hide
End Sub
