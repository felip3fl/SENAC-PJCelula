VERSION 5.00
Begin VB.Form frmFornecedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedor - CELL SOFT"
   ClientHeight    =   7230
   ClientLeft      =   8175
   ClientTop       =   3915
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmFornecedor.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdDescricao 
      Height          =   495
      Left            =   6120
      Picture         =   "frmFornecedor.frx":30FC2
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Inserir uma descrição para esse funcionario"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FCF0E7&
      Height          =   1845
      Left            =   240
      TabIndex        =   38
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   7320
      Width           =   7215
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmFornecedor.frx":3527D
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
      Picture         =   "frmFornecedor.frx":38437
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exibir funcionario Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmFornecedor.frx":3B61E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmFornecedor.frx":40E9D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Excluir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmFornecedor.frx":457EB
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
      Picture         =   "frmFornecedor.frx":4A736
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Editar um Registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   6120
      Picture         =   "frmFornecedor.frx":4F77C
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   27
      ToolTipText     =   "Foto do Funcionario"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmFornecedor.frx":5528C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Excluir a foto desse funcionario"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   25
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtEndereco 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   24
      ToolTipText     =   "Endereço do Funcionario - Max.: 50 Caracter"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox txtRG 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   23
      ToolTipText     =   "RG do Funcionario - Max.: 10 Caracter"
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox txtTelefone 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      ToolTipText     =   "Numero do Celular do Funcionario - Max.: 14 Caracter"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtData_Nasc 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   21
      ToolTipText     =   "Data de Nascimento do Funcionario - Padrão: 01/01/2000"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtCelular 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      ToolTipText     =   "Numero do Celular do Funcionario - Max.:  14 Caracter"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNova_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmFornecedor.frx":595B0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Inserir nova foto para esse funcionario"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPesquisa 
      Height          =   615
      Left            =   240
      Picture         =   "frmFornecedor.frx":5D7AC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pesquisar o Codigo, Cargo ou Nome do funcionario"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtCargo_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Digite o Cargo de um funcionario para a pesquisa"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtCod_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Digite o Codigo de um Funcionario para a pesquisa"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtNome_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Digite um Nome de um funcionario para a pesquisa"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmFornecedor.frx":61923
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmFornecedor.frx":65EB7
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exibir o Ultimo Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmFornecedor.frx":69236
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exibir o Primeiro Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
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
      TabIndex        =   37
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor:"
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
      TabIndex        =   36
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do  Fornecedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   35
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
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
      Left            =   4200
      TabIndex        =   33
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
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
      TabIndex        =   32
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Emdereço da Empresa"
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
      TabIndex        =   31
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo do"
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
      TabIndex        =   30
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CMPJ da Empresa:"
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
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Fornecedor"
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
      TabIndex        =   28
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor:"
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
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo do"
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
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Fornecedor:"
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
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DA8145&
      X1              =   2160
      X2              =   2160
      Y1              =   3480
      Y2              =   6960
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade da Empresa:"
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
      Top             =   5640
      Width           =   2055
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
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

