VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_listar 
   Caption         =   "Listar stock"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Bindings        =   "frm_listar.frx":0000
      Height          =   2175
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.CommandButton cmb_volver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Data Data_listar 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Mis documentos\Descargas\Proyecto\stock.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "productos"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frm_listar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_volver_Click()
Me.Hide
End Sub

Private Sub Form_Activate()



End Sub

