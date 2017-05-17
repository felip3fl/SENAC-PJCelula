VERSION 5.00
Begin VB.Form frmFuncionario 
   BackColor       =   &H00FCF0E7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadasdro de Funcionarios - CELL SOFT"
   ClientHeight    =   7245
   ClientLeft      =   7980
   ClientTop       =   3750
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCelula.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdDescricao 
      Height          =   495
      Left            =   6120
      Picture         =   "frmCelula.frx":30FC2
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Inserir uma descrição para esse funcionario"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FCF0E7&
      Height          =   1845
      Left            =   240
      TabIndex        =   50
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   7320
      Width           =   7215
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmCelula.frx":3527D
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exibir funcionario Proximo"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmCelula.frx":38437
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exibir funcionario Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmCelula.frx":3B61E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmCelula.frx":40E9D
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Excluir este registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmCelula.frx":457EB
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir este registro"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmCelula.frx":4A736
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmCelula.frx":503D6
      Style           =   1  'Graphical
      TabIndex        =   48
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
      Picture         =   "frmCelula.frx":54CF2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Editar um Registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmCelula.frx":59D38
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdNova_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmCelula.frx":5E2CC
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Inserir nova foto para esse funcionario"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPesquisa 
      Height          =   615
      Left            =   240
      Picture         =   "frmCelula.frx":624C8
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Pesquisar o Codigo, Cargo ou Nome do funcionario"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtCargo_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Digite o Cargo de um funcionario para a pesquisa"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtCod_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Digite o Codigo de um Funcionario para a pesquisa"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtNome_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   26
      ToolTipText     =   "Digite um Nome de um funcionario para a pesquisa"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmCelula.frx":6663F
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exibir o Primeiro Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmCelula.frx":69A0E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exibir o Ultimo Funcionario"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtSexo 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Sexo do Funcionario - Max.: 9 Caracter"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Codigo do Funcionario"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtCelular 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      ToolTipText     =   "Numero do Celular do Funcionario - Max.:  14 Caracter"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtData_Nasc 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Data de Nascimento do Funcionario - Padrão: 01/01/2000"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtCidade 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "Cidade do Funcionario - Max.: 20 Caracter"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtTelefone 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "Numero do Celular do Funcionario - Max.: 14 Caracter"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtCargo 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Cargo do Funcionario - Max.: 20 Caracter"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtRG 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "RG do Funcionario - Max.: 10 Caracter"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtCPF 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "CPF do Funcionario - Max.: 11 Caracter"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtCEP 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "Numero do CEP do Funcionario - Max.: 9 Caracter"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtEndereco 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      ToolTipText     =   "Endereço do Funcionario - Max.: 50 Caracter"
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton cmdExcluir_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmCelula.frx":6CD8D
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Excluir a foto desse funcionario"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   6120
      Picture         =   "frmCelula.frx":710B1
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   0
      ToolTipText     =   "Foto do Funcionario"
      Top             =   1800
      Width           =   1335
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2400
      TabIndex        =   47
      Top             =   1320
      Width           =   5055
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
      TabIndex        =   46
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
      TabIndex        =   45
      Top             =   3960
      Width           =   1095
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
      TabIndex        =   44
      Top             =   4920
      Width           =   615
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
      TabIndex        =   43
      Top             =   5640
      Width           =   1335
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
      TabIndex        =   42
      Top             =   3480
      Width           =   975
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
      TabIndex        =   41
      Top             =   3480
      Width           =   1815
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
      TabIndex        =   40
      Top             =   1800
      Width           =   1095
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
      TabIndex        =   39
      Top             =   2040
      Width           =   615
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
      TabIndex        =   38
      Top             =   4200
      Width           =   255
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
      TabIndex        =   37
      Top             =   4200
      Width           =   375
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
      TabIndex        =   36
      Top             =   4920
      Width           =   735
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
      TabIndex        =   35
      Top             =   4920
      Width           =   855
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
      TabIndex        =   34
      Top             =   3480
      Width           =   975
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
      TabIndex        =   33
      Top             =   6360
      Width           =   495
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
      TabIndex        =   32
      Top             =   6360
      Width           =   735
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
      TabIndex        =   31
      Top             =   5640
      Width           =   3015
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
      TabIndex        =   30
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label2 
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
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
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
      TabIndex        =   28
      Top             =   6960
      Width           =   3015
   End
End
Attribute VB_Name = "frmFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Controle As Boolean
    Dim varFunCod As String


Private Sub cmdDescricao_Click()
    frmFuncionario.Height = 9390
End Sub


Private Sub txtDescricao_LostFocus()
    frmFuncionario.Height = 7575
End Sub

Private Sub Form_Load()
    AbrirDB
    
    Bloquear
    ExibirDados
    cmdSalvar.Visible = False ' esconde o botão salvar, atrás do formulário
    cmdCancelar.Visible = False
End Sub
    
Private Sub Bloquear()
    txtNome.Locked = True
    'txtCodigo.Locked = True
    txtEndereco.Locked = True
    txtCidade.Locked = True
    txtCEP.Locked = True
    txtData_Nasc.Locked = True
    txtTelefone.Locked = True
    txtCelular.Locked = True
    txtCPF.Locked = True
    txtRG.Locked = True
    txtCargo.Locked = True
End Sub

Private Sub Desbloquear()
    txtNome.Locked = False
    'txtCodigo.Locked = False
    txtEndereco.Locked = False
    txtCidade.Locked = False
    txtCEP.Locked = False
    txtData_Nasc.Locked = False
    txtTelefone.Locked = False
    txtCelular.Locked = False
    txtCPF.Locked = False
    txtRG.Locked = False
    txtCargo.Locked = False
End Sub

Private Sub ExibirDados()
    Limpar
    On Error Resume Next
    txtNome.Text = tblFuncionario!nome
    txtCodigo.Text = tblFuncionario!codigo
    txtEndereco.Text = tblFuncionario!endereco
    txtCidade.Text = tblFuncionario!cidade
    txtCEP.Text = tblFuncionario!cep
    txtData_Nasc.Text = tblFuncionario!data_nasc
    txtTelefone.Text = tblFuncionario!telefone
    txtCelular.Text = tblFuncionario!celular
    txtCPF.Text = tblFuncionario!cpf
    txtRG.Text = tblFuncionario!rg
    txtCargo.Text = tblFuncionario!cargo
    txtSexo.Text = tblFuncionario!sexo
    On Error GoTo 0
End Sub

Private Sub cmdNovo_Click()
    Desbloquear
    Limpar
    txtNome.BackColor = &HFFFFFF
    txtEndereco.BackColor = &HFFFFFF
    txtData_Nasc.BackColor = &HFFFFFF
    txtRG.BackColor = &HFFFFFF
    txtCPF.BackColor = &HFFFFFF
    txtCelular.BackColor = &HFFFFFF
    txtTelefone.BackColor = &HFFFFFF
    txtEndereco.BackColor = &HFFFFFF
    txtCidade.BackColor = &HFFFFFF
    txtCEP.BackColor = &HFFFFFF
    txtCargo.BackColor = &HFFFFFF
    txtSexo.BackColor = &HFFFFFF
    
    cmdNovo.Visible = False ' esconde o frame quando clica em NOVO
    cmdAlterar.Visible = False
    cmdImprimir.Visible = False
    cmdExcluir.Visible = False
    cmdSalvar.Visible = True ' mostra o SALVAR que estava escondido atrás da FRAME
    cmdCancelar.Visible = True
    txtCodigo.Text = Format(Val(varFunCod) + 1, "00000")
    txtNome.SetFocus
    Controle = True
End Sub

    Private Sub cmdAlterar_Click()
    Desbloquear
    
    txtNome.BackColor = &HFFFFFF
    txtEndereco.BackColor = &HFFFFFF
    txtData_Nasc.BackColor = &HFFFFFF
    txtRG.BackColor = &HFFFFFF
    txtCPF.BackColor = &HFFFFFF
    txtCelular.BackColor = &HFFFFFF
    txtTelefone.BackColor = &HFFFFFF
    txtEndereco.BackColor = &HFFFFFF
    txtCidade.BackColor = &HFFFFFF
    txtCEP.BackColor = &HFFFFFF
    txtCargo.BackColor = &HFFFFFF
    txtSexo.BackColor = &HFFFFFF
    
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    cmdImprimir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
    txtNome.SetFocus
    Controle = False
    
End Sub
    Private Sub AbrirDB()
    Set ACelula = DBEngine.Workspaces(0) ' (0) significa índice, primeiro
    On Error GoTo ErroAbrir ' caso dê erro...desvia, não continua nas linhas debaixo
    ACelula.BeginTrans ' ...inicia a transação do BD
    Set BCelula = ACelula.OpenDatabase(App.Path & "\Celula.mdb", False) ' bagenda(área) recebe aagenda(tabela) salvo no local determinado
    Set tblFuncionario = BCelula.OpenRecordset("tblfuncionario", dbOpenDynaset) ' abre como conjunto de registros, dbopenDynaset( salvar, exluir, alterar, etc)
    Set TbCelula = tblFuncionario.OpenRecordset() ' abre a tabela física no local determinado tblagenda.openrecordset
    'Aqui
    Set tblCodigos = BCelula.OpenRecordset("tblcodigos", dbOpenDynaset)
    Set TbCodigos = tblCodigos.OpenRecordset()
    varFunCod = tblCodigos!funcod
    
    ACelula.CommitTrans ' encerra a transação do Banco de Dados
    On Error GoTo 0 ' caso dê erro, sai da rotina
    Exit Sub ' encerra a rotina
ErroAbrir:
    Dim Aviso As Integer
    Aviso = MsgBox("Erro ao acessar os Dados!", vbCritical, "Aviso!!!")
    If Aviso = vbOK Then ' caso clique no OK...
        ACelula.Rollback ' rollback suspende a transação do Banco de Dados
        On Error GoTo 0 ' Não mostra o erro
        Exit Sub
    End If
    Resume ' Mostra o erro
End Sub


Private Sub cmdSair_Click()
    'Aqui
    ACelula.BeginTrans
    tblCodigos.Edit
    tblCodigos!funcod = varFunCod
    tblCodigos.Update
    ACelula.CommitTrans
    ACelula.Close ' fecha o banco de dados
    Unload Me
End Sub


Private Sub cmdSalvar_Click()
    ACelula.BeginTrans ' inicia transação com o banco de dados
    If Controle = True Then
        tblFuncionario.AddNew ' adiciona novo registro em branco
        tblFuncionario!nome = txtNome.Text
        tblFuncionario!codigo = txtCodigo.Text
        tblFuncionario!endereco = txtEndereco.Text
        tblFuncionario!cidade = txtCidade.Text
        tblFuncionario!data_nasc = txtData_Nasc.Text
        tblFuncionario!telefone = txtTelefone.Text
        tblFuncionario!celular = txtCelular.Text
        tblFuncionario!cep = txtCEP.Text
        tblFuncionario!sexo = txtSexo.Text
        tblFuncionario!cpf = txtCPF.Text
        tblFuncionario!rg = txtRG.Text
        tblFuncionario!cargo = txtCargo.Text
        tblFuncionario.Update ' guarda as informações, atualiza
        MsgBox " Inclusão com sucesso!", vbInformation, " Funcionario"
        'Aqui
        varFunCod = Val(varFunCod) + 1
    Else
        tblFuncionario.Edit ' edita o registro mostrado
        tblFuncionario!nome = txtNome.Text
        tblFuncionario!codigo = txtCodigo.Text
        tblFuncionario!endereco = txtEndereco.Text
        tblFuncionario!cidade = txtCidade.Text
        tblFuncionario!data_nasc = txtData_Nasc.Text
        tblFuncionario!telefone = txtTelefone.Text
        tblFuncionario!celular = txtCelular.Text
        tblFuncionario!cep = txtCEP.Text
        tblFuncionario!sexo = txtSexo.Text
        tblFuncionario!cpf = txtCPF.Text
        tblFuncionario!rg = txtRG.Text
        tblFuncionario!cargo = txtCargo.Text
        tblFuncionario.Update ' guarda as informações, atualiza
        MsgBox " Alterado com sucesso!", vbInformation, " Funcionario"
    End If
        ACelula.CommitTrans ' Encerra a transação do banco de dados
        
    txtNome.BackColor = &HFCF0E7
    txtEndereco.BackColor = &HFCF0E7
    txtData_Nasc.BackColor = &HFCF0E7
    txtRG.BackColor = &HFCF0E7
    txtCPF.BackColor = &HFCF0E7
    txtCelular.BackColor = &HFCF0E7
    txtTelefone.BackColor = &HFCF0E7
    txtEndereco.BackColor = &HFCF0E7
    txtCidade.BackColor = &HFCF0E7
    txtCEP.BackColor = &HFCF0E7
    txtCargo.BackColor = &HFCF0E7
    txtSexo.BackColor = &HFCF0E7
        
    Bloquear
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    cmdImprimir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    Controle = True ' Deixa habilitado pois será inserido novo registro
    
End Sub
Private Sub cmdCancelar_Click()
    
    txtNome.BackColor = &HFCF0E7
    txtEndereco.BackColor = &HFCF0E7
    txtData_Nasc.BackColor = &HFCF0E7
    txtRG.BackColor = &HFCF0E7
    txtCPF.BackColor = &HFCF0E7
    txtCelular.BackColor = &HFCF0E7
    txtTelefone.BackColor = &HFCF0E7
    txtEndereco.BackColor = &HFCF0E7
    txtCidade.BackColor = &HFCF0E7
    txtCEP.BackColor = &HFCF0E7
    txtCargo.BackColor = &HFCF0E7
    txtSexo.BackColor = &HFCF0E7
    
    Bloquear
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdImprimir.Visible = True
    cmdExcluir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    ExibirDados
End Sub

Private Sub cmdExcluir_Click()
    Dim Excluir As String
    Excluir = MsgBox("Deseja excluir este registro?", vbQuestion + vbYesNo, "Funcionário")
    If Excluir <> vbYes Then
        Exit Sub
    Else
        If Not tblFuncionario.EOF = True And Not tblFuncionario.BOF = True Then ' BOF: começo de arquivo, EOF: fim de arquivo
            ACelula.BeginTrans  ' inicia transação do banco de dados, para efetuar a ação.
            tblFuncionario.Delete ' deleta o registro
            tblFuncionario.MovePrevious ' move para o registro anterior
            ExibirDados ' exibe o dado anterior
            ACelula.CommitTrans ' Encerra a transação do banco de dados
         Else ' se estiver vazio o banco de dados:
            Excluir = MsgBox("Fim de arquivo encontrado!", vbInformation, "Funcionário")
         End If
    End If
End Sub


Private Sub cmdPrimeiro_Click()
    If tblFuncionario.BOF = True And tblFuncionario.EOF = True Then
        MsgBox "O registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Primeiro ' Se ocorrer erro, ele vai para o primeiro registro
    tblFuncionario.MoveFirst ' vai para o primeiro registro da agenda
    ExibirDados
    Exit Sub
Primeiro:
    If Err.Number = 3021 Then ' se clicar no BOF apresenta esse ero
        MsgBox "Você já está no primeiro registro!", vbInformation, "Aviso!"
        tblFuncionario.MoveFirst ' mostra o primeiro registro para não ficar em branco a tabela
        ExibirDados
        Exit Sub
    End If
    
End Sub


Private Sub cmdUltimo_Click()
    If tblFuncionario.BOF = True And tblFuncionario.EOF = True Then
        MsgBox "O registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Ultimo ' Se ocorrer erro, ele vai para o primeiro registro
    tblFuncionario.MoveLast ' vai para o último registro da agenda
    ExibirDados
    Exit Sub
Ultimo:
    If Err.Number = 3021 Then ' se clicar no BOF apresenta esse ero
        MsgBox "Você já está no último registro!", vbInformation, "Aviso!"
        tblFuncionario.MoveLast ' mostra o último registro para não ficar em branco a tabela
        ExibirDados ' exibe o último dado
        Exit Sub
    End If
End Sub
' Rotina para retroceder um registro
Private Sub cmdAnterior_Click()
    If tblFuncionario.BOF = True And tblFuncionario.EOF = True Then
        MsgBox "O registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Anterior ' Se ocorrer erro, ele vai para o primeiro registro
    tblFuncionario.MovePrevious ' vai para um registro anterior da agenda
    If tblFuncionario.BOF = False Then
        ExibirDados ' exibe o dado do registro anterior
    End If
    Exit Sub
Anterior:
    If Err.Number = 3021 Then ' se clicar no BOF apresenta esse ero
        MsgBox "Você já está no primeiro registro!", vbInformation, "Aviso!"
        tblFuncionario.MoveFirst ' move o registro primeiro para não ficar em branco a tabela
        ExibirDados ' exibe o dado anterior
        Exit Sub
    End If
End Sub
' Rotina para avançar um regitro
Private Sub cmdProximo_Click()
    If tblFuncionario.BOF = True And tblFuncionario.EOF = True Then
        MsgBox "Agenda está vazia", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Proximo ' Se ocorrer erro, ele vai para o primeiro registro
    tblFuncionario.MoveNext ' vai para um proximo registro da agenda
    If tblFuncionario.EOF = False Then ' se estiver no último registro
        ExibirDados ' exibe o dado do próximo registro
    End If
    Exit Sub
Proximo:
    If Err.Number = 3021 Then ' se clicar no BOF apresenta esse ero
        MsgBox "Você já está no último registro!", vbInformation, "Aviso!"
        tblFuncionario.MoveLast ' move o registro final para não ficar em branco a tabela
        ExibirDados ' exibe o dado anterior
        Exit Sub
    End If
End Sub

' Rotina para pesquisar cadastro
Private Sub cmdPesquisa_Click()
    If Not txtNome_Pesq.Text = "" Then ' Se não estiver em branco
        tblFuncionario.FindFirst "nome like '" + txtNome_Pesq.Text + "*'" ' pesquisar na tabela agenda pelo primeiro registro, o conteúdo que digitou no txtnomepesq (*= busca o primeiro nome, caso digite o primeiro nome, se não colocar, precisa colocar o nome completo)
        ExibirDados
        If tblFuncionario.NoMatch Then ' se não for localizado
            MsgBox "O nome " + txtNome_Pesq.Text + " não foi localizado.", vbInformation, "Aviso!"
            txtNome_Pesq.Text = "" ' limpa o campo
            txtNome_Pesq.SetFocus ' o cursor fica em txtnomepesq.text
        End If
        txtNome_Pesq.Text = Empty
        txtCod_Pesq.Text = Empty
    Else
        If Not txtCod_Pesq.Text = "" Then
            txtCod_Pesq.Text = Format(txtCod_Pesq.Text, "00000")
            tblFuncionario.FindFirst "codigo like '" + txtCod_Pesq.Text + "'" ' ' pesquisar na tabela agenda pelo primeiro registro, o conteúdo que digitou no txtnomepesq (*= busca o primeiro nome, caso digite o primeiro nome, se não colocar, precisa colocar o nome completo)
            ExibirDados
            If tblFuncionario.NoMatch Then
                MsgBox "O codigo " + txtCod_Pesq.Text + " não foi localizado.", vbInformation, "Aviso!"
                txtCod_Pesq.Text = ""
                txtCod_Pesq.SetFocus
            End If
            txtCargo_Pesq.Text = ""
            txtCargo_Pesq.SetFocus
        Else
            tblFuncionario.FindFirst " cargo like '" + txtCargo_Pesq.Text + "*'"
            ExibirDados
            If tblFuncionario.NoMatch Then
                MsgBox " O Cargo " + txtCargo_Pesq + " não foi localizado.", vbInformation, "Aviso!"
                txtCargo_Pesq.Text = ""
                txtCargo_Pesq.SetFocus
             End If
        End If
    End If
    txtNome_Pesq.Text = ""
    txtCod_Pesq.Text = ""
End Sub



' Utilzar ENTER ao invés de TAB
'Private Sub txtCargo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       SendKeys "{TAB}"
'        End If


'
' Rotina de consistência da data de aniversário
Private Sub txtData_nasc_LostFocus()
    If Not IsDate(txtData_Nasc.Text) Then ' Se a data não for da forma de data...
        MsgBox "Data INVÁLIDA!", vbInformation, "Funcionário"
        txtData_Nasc.Text = ""
        txtData_Nasc.SetFocus
    End If
End Sub


Private Sub Limpar()
    txtNome.Text = ""
    txtCodigo.Text = ""
    txtEndereco.Text = ""
    txtData_Nasc.Text = ""
    txtCidade.Text = ""
    txtCEP.Text = ""
    txtTelefone.Text = ""
    txtCelular.Text = ""
    txtCPF.Text = ""
    txtRG.Text = ""
    txtCargo.Text = ""
End Sub

Private Sub txtTelefonePesq_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 8)) Then
        KeyAscii = 0
    End If
End Sub




