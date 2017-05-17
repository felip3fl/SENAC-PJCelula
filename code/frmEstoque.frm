VERSION 5.00
Begin VB.Form frmEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " - CELL SOFT"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEstoque.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   6120
      Picture         =   "frmEstoque.frx":22172
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   34
      ToolTipText     =   "Foto do Funcionario"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmEstoque.frx":27C82
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Excluir a foto desse funcionario"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   32
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtEndereco 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      ToolTipText     =   "Endereço do Funcionario - Max.: 50 Caracter"
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtCEP 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   30
      ToolTipText     =   "Numero do CEP do Funcionario - Max.: 9 Caracter"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtCPF 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   29
      ToolTipText     =   "CPF do Funcionario - Max.: 11 Caracter"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtRG 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      ToolTipText     =   "RG do Funcionario - Max.: 10 Caracter"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtCargo 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   3840
      TabIndex        =   27
      ToolTipText     =   "Cargo do Funcionario - Max.: 20 Caracter"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtTelefone 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   26
      ToolTipText     =   "Numero do Celular do Funcionario - Max.: 14 Caracter"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtCidade 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   25
      ToolTipText     =   "Cidade do Funcionario - Max.: 20 Caracter"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtData_Nasc 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   24
      ToolTipText     =   "Data de Nascimento do Funcionario - Padrão: 01/01/2000"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtCelular 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   23
      ToolTipText     =   "Numero do Celular do Funcionario - Max.:  14 Caracter"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtSexo 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4440
      TabIndex        =   21
      ToolTipText     =   "Sexo do Funcionario - Max.: 9 Caracter"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmEstoque.frx":2BFA6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exibir o Ultimo Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmEstoque.frx":2F325
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exibir o Primeiro Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdNova_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmEstoque.frx":326F4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Inserir nova foto para esse funcionario"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdDescricao 
      Height          =   495
      Left            =   6120
      Picture         =   "frmEstoque.frx":368F0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Inserir uma descrição para esse funcionario"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmEstoque.frx":3ABAB
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exibir funcionario Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmEstoque.frx":3DD92
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exibir funcionario Proximo"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtNome_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Digite um Nome de um funcionario para a pesquisa"
      Top             =   5880
      Width           =   1695
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
      Picture         =   "frmEstoque.frx":40F4C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pesquisar o Codigo, Cargo ou Nome do funcionario"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmEstoque.frx":450C3
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
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
      Picture         =   "frmEstoque.frx":49657
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Editar um Registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmEstoque.frx":4E69D
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmEstoque.frx":535E8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Excluir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmEstoque.frx":57F36
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
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
      TabIndex        =   49
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionario:"
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
      TabIndex        =   48
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
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
      TabIndex        =   47
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
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
      TabIndex        =   46
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label13 
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
      TabIndex        =   45
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
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
      TabIndex        =   44
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexo:"
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
      Left            =   4440
      TabIndex        =   43
      Top             =   3480
      Width           =   975
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
      TabIndex        =   42
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
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
      TabIndex        =   41
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
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
      TabIndex        =   40
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "RG:"
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
      TabIndex        =   39
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo:"
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
      Left            =   3840
      TabIndex        =   38
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
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
      TabIndex        =   37
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Nascimento:"
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
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cadatro"
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
      TabIndex        =   35
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
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Funcionarios"
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
      Height          =   1935
      Left            =   2400
      TabIndex        =   14
      Top             =   1320
      Width           =   5055
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
      Caption         =   "Nome:"
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
      Width           =   1335
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
      Caption         =   "Cargo:"
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
      Width           =   615
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
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionario:"
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
Attribute VB_Name = "frmEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
