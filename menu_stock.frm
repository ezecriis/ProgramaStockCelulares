VERSION 5.00
Begin VB.Form menu_stock 
   Caption         =   "Stock"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   7425
   LinkTopic       =   "Form3"
   ScaleHeight     =   4785
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnu_productos 
      Caption         =   "Productos"
      Begin VB.Menu mnu_agregar_prod 
         Caption         =   "Agregar producto"
      End
      Begin VB.Menu mnu_modificar_prod 
         Caption         =   "Modificar producto"
      End
      Begin VB.Menu mnu_eliminar_prod 
         Caption         =   "Eliminar producto"
      End
   End
   Begin VB.Menu mnu_buscar 
      Caption         =   "Buscar.."
      Begin VB.Menu mnu_buscar_cod 
         Caption         =   "Por codigo"
      End
      Begin VB.Menu mnu_buscar_prod 
         Caption         =   "Por producto"
      End
   End
   Begin VB.Menu mnu_listado 
      Caption         =   "Listado"
   End
   Begin VB.Menu mnu_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "menu_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_agregar_prod_Click()
frm_agregar_celular.Show
End Sub

Private Sub mnu_buscar_cod_Click()
buscar_nombre.Show
End Sub

Private Sub mnu_buscar_prod_Click()
buscar_modelo.Show
End Sub

Private Sub mnu_eliminar_prod_Click()
frm_eliminar_prod.Show
End Sub

Private Sub mnu_listado_Click()
frm_listado.Show
End Sub

Private Sub mnu_modificar_prod_Click()
frm_modificar_prod.Show
End Sub

Private Sub mnu_salir_Click()
Me.Hide
End Sub
