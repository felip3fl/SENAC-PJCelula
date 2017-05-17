VERSION 5.00
Begin VB.Form frmVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas - CELL SOFT"
   ClientHeight    =   7230
   ClientLeft      =   8340
   ClientTop       =   4065
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmVendas.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   7695
   Begin VB.TextBox Text3 
      BackColor       =   &H00FCF0E7&
      Height          =   1845
      Left            =   240
      TabIndex        =   30
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   7320
      Width           =   7215
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmVendas.frx":30FC2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exibir funcionario Proximo"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmVendas.frx":3417C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exibir funcionario Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmVendas.frx":37363
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmVendas.frx":3CBE2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Excluir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmVendas.frx":41530
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlterar 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmVendas.frx":4647B
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Editar um Registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtValor_Total 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      ToolTipText     =   "Cargo do Funcionario - Max.: 20 Caracter"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   6120
      Picture         =   "frmVendas.frx":4B4C1
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   20
      ToolTipText     =   "Foto do Funcionario"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtData_Compra 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Cargo do Funcionario - Max.: 20 Caracter"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txtCod_Venda 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmVendas.frx":52DA7
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exibir o Ultimo Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmVendas.frx":56126
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exibir o Primeiro Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtData_Compra_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Digite um Nome de um funcionario para a pesquisa"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtCod_Venda_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Digite o Codigo de um Funcionario para a pesquisa"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtCod_Func_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Digite o Cargo de um funcionario para a pesquisa"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdPesquisa 
      Height          =   615
      Left            =   240
      Picture         =   "frmVendas.frx":594F5
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pesquisar o Codigo, Cargo ou Nome do funcionario"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmVendas.frx":5D66C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmVendas.frx":61C00
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmVendas.frx":678A0
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total da Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo do Funcionario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "desenvolved by Group Célula Supermercados © 2009"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo da Venda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vendas"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00DA8145&
      X1              =   7440
      X2              =   6120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Compra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DA8145&
      X1              =   2160
      X2              =   2160
      Y1              =   3480
      Y2              =   6960
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Funcionario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "da Venda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Controle As Boolean
    Dim varFunCod As String
  
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Bloquear()

    txtCod_Venda.Locked = True
    txtCodigo.Locked = True
    txtData_Compra.Locked = True
    txtValor_Total.Locked = True
    
  
End Sub

Private Sub Desbloquear()

    txtCodigo.Locked = False
    txtData_Compra.Locked = False
    txtValor_Total.Locked = False
    
  
    
End Sub

Private Sub cmdNovo_Click()

    Desbloquear
    Limpar
    
    txtCodigo.BackColor = &HFFFFFF
    txtData_Compra.BackColor = &HFFFFFF
    txtValor_Total.BackColor = &HFFFFFF
    
    
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdImprimir.Visible = False
    cmdExcluir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
    
    'txtCod_Produto.Text = Format(Val(varFunCod) + 1, "00000")
    
    
    Controle = True
    
End Sub

    Private Sub cmdAlterar_Click()
    Desbloquear
    
    
    txtCodigo.Locked = True
    txtData_Compra.BackColor = &HFFFFFF
    txtValor_Total.BackColor = &HFFFFFF
    
    
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdImprimir.Visible = False
    cmdExcluir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
   'txtCod_Produto.Text = Format(Val(varFunCod) + 1, "00000")
   
    Controle = True
    
End Sub

Private Sub Limpar()
    
    txtCod_Venda.Text = ""
    txtCodigo.Text = ""
    txtData_Compra.Text = ""
    txtValor_Total.Text = ""
End Sub

Private Sub cmdCancelar_Click()
    
   
    txtCodigo.BackColor = &HFCF0E7
    txtData_Compra.BackColor = &HFCF0E7
    txtValor_Total.BackColor = &HFCF0E7
    
    Bloquear
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdImprimir.Visible = True
    cmdExcluir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    
End Sub

Private Sub Form_Load()
    Bloquear
    cmdSalvar.Visible = False
    cmdCancelar.Visible = True
End Sub



