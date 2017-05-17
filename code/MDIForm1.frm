VERSION 5.00
Begin VB.MDIForm MDIcell 
   BackColor       =   &H8000000C&
   Caption         =   "CELL SOFT"
   ClientHeight    =   7335
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12300
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Abrir Formularios"
      Begin VB.Menu mnuFornecedor 
         Caption         =   "&Fornecedor"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFuncionario 
         Caption         =   "&Funcionario"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuItensdaVendas 
         Caption         =   "&Itens da Vendas"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuVendas 
         Caption         =   "&Vendas"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuProduto 
         Caption         =   "&Produto"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuseparar 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&Sobre o Software"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "MDIcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFornecedor_Click()
    Load frmFornecedor
    frmFornecedor.Show
End Sub

Private Sub mnuFuncionario_Click()
    Load frmFuncionario
    frmFuncionario.Show
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuItensdaVendas_Click()
    Load frmItensdaVendas
    frmItensdaVendas.Show
End Sub

Private Sub mnuProduto_Click()
    Load frmProduto
    frmProduto.Show
End Sub

Private Sub mnuSair_Click()
    End
End Sub



Private Sub mnuSobre_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuVendas_Click()
    Load frmVendas
    frmVendas.Show
End Sub


