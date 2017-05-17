VERSION 5.00
Begin VB.Form frmItensdaVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Itens da Vendas - CELL SOFT"
   ClientHeight    =   7215
   ClientLeft      =   8310
   ClientTop       =   4110
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmItensdaVendas.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdDescricao 
      Height          =   495
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":30FC2
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Inserir uma descrição para esse funcionario"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FCF0E7&
      Height          =   1845
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   7320
      Width           =   7215
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmItensdaVendas.frx":3527D
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exibir funcionario Proximo"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":38437
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exibir funcionario Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":3B61E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmItensdaVendas.frx":40E9D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Excluir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmItensdaVendas.frx":457EB
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
      Picture         =   "frmItensdaVendas.frx":4A736
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
      Picture         =   "frmItensdaVendas.frx":4F77C
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   21
      ToolTipText     =   "Foto do Funcionario"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":5528C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Excluir a foto desse funcionario"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtRG 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   18
      ToolTipText     =   "RG do Funcionario - Max.: 10 Caracter"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtCelular 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   17
      ToolTipText     =   "Numero do Celular do Funcionario - Max.:  14 Caracter"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmItensdaVendas.frx":595B0
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exibir o Ultimo Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":5C92F
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exibir o Primeiro Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdNova_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmItensdaVendas.frx":5FCFE
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Inserir nova foto para esse funcionario"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtCod_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Digite o Codigo de um Funcionario para a pesquisa"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtCargo_Pesq 
      BackColor       =   &H8000000A&
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
      Picture         =   "frmItensdaVendas.frx":63EFA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pesquisar o Codigo, Cargo ou Nome do funcionario"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmItensdaVendas.frx":68071
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nenhum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   33
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nenhum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Item de Venda"
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
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label5 
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
      Left            =   2400
      TabIndex        =   29
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "do Produto:"
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
      TabIndex        =   28
      Top             =   3120
      Width           =   2175
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
      TabIndex        =   27
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item da Venda:"
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
      TabIndex        =   26
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço do Produto:"
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
      TabIndex        =   25
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Quatidade de"
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
      Top             =   3960
      Width           =   2895
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
      TabIndex        =   23
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Itens da Vendas"
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
      TabIndex        =   22
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
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DA8145&
      X1              =   2160
      X2              =   2160
      Y1              =   3480
      Y2              =   6240
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
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Item da Venda:"
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
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "frmItensdaVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub
