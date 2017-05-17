VERSION 5.00
Begin VB.Form frmProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos - CELL SOFT"
   ClientHeight    =   7230
   ClientLeft      =   8340
   ClientTop       =   4140
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmProduto.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdDescricao 
      Height          =   495
      Left            =   6120
      Picture         =   "frmProduto.frx":30FC2
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Inserir uma descrição para esse funcionario"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtNome_Fornecedor 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   40
      ToolTipText     =   "Codigo do Fornecedor"
      Top             =   6120
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FCF0E7&
      Height          =   1845
      Left            =   240
      TabIndex        =   37
      ToolTipText     =   "Nome do Funcionario - Max.: 50 Caracter"
      Top             =   7320
      Width           =   7215
   End
   Begin VB.CommandButton cmdProximo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmProduto.frx":3527D
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exibir o  Proximo Produto"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdAnterior 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmProduto.frx":38437
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exibir o Produto Anterior"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   6120
      Picture         =   "frmProduto.frx":3B61E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sair desse Programa"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmProduto.frx":40E9D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Excluir este registro"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmProduto.frx":457EB
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir este registro"
      Top             =   240
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
      Picture         =   "frmProduto.frx":4A736
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Editar um Registro"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtQuant_Estoque 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   30
      ToolTipText     =   "Quantidade do Estoque"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtCod_Fornecedor 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   29
      ToolTipText     =   "Codigo do Fornecedor"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtPreco_Produto 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   28
      ToolTipText     =   "Preço do Produto"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtCod_Forn_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "Digite um Nome de um Produto para a pesquisa"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtCod_Prod_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Digite o Codigo de um  Produto para a pesquisa"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtNome_Prod_Pesq 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "Digite o Cargo de um  Produto para a pesquisa"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdPesquisa 
      Height          =   615
      Left            =   240
      Picture         =   "frmProduto.frx":4F77C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Pesquisar Codigo do Produto, Nome do produto ou Fornecedor do Produto"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   6120
      Picture         =   "frmProduto.frx":538F3
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   17
      ToolTipText     =   "Foto do Produto"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmProduto.frx":59403
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Excluir a foto desse Produto"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdUltimo 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6840
      Picture         =   "frmProduto.frx":5D727
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exibir o Ultimo  Produto"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdPrimeiro 
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   6120
      Picture         =   "frmProduto.frx":60AA6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exibir o Primeiro  Produto"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdNova_Foto 
      Height          =   495
      Left            =   6120
      Picture         =   "frmProduto.frx":63E75
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Inserir nova foto para esse Produto"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtDescricao_Produto 
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Nome do Produto"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtCod_Produto 
      Alignment       =   2  'Center
      BackColor       =   &H00FCF0E7&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Codigo do Produto"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovo 
      Height          =   1215
      Left            =   0
      Picture         =   "frmProduto.frx":68071
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Criar um novo Registro"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   1215
      Left            =   4560
      Picture         =   "frmProduto.frx":6C605
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmProduto.frx":722A5
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço"
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
      TabIndex        =   42
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Fornecedor"
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
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "do Estoque"
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
      TabIndex        =   36
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label11 
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
      TabIndex        =   35
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "do Produto"
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
      TabIndex        =   34
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
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
      Left            =   2640
      TabIndex        =   33
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label7 
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
      TabIndex        =   32
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do Produto"
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
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label5 
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
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Formecedor:"
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
      TabIndex        =   26
      Top             =   5640
      Width           =   1815
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
      Caption         =   "Nome do Produto:"
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
      TabIndex        =   25
      Top             =   4920
      Width           =   1815
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
      TabIndex        =   24
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      TabIndex        =   23
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "desenvolved by Group Célula Supermercados © 2009"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00DA8145&
      X1              =   7440
      X2              =   6120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
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
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
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
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Codiso do"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Produtos"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   6
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Controle As Boolean
    Dim varProCod As String
    
    
Private Sub cmdSair_Click()
    ACelula.BeginTrans
    tblCodigos.Edit
    tblCodigos!procod = varProCod
    tblCodigos.Update
    ACelula.CommitTrans
    ACelula.Close
    Unload Me
End Sub





Private Sub Bloquear()

    txtCod_Produto.Locked = True
    txtDescricao_Produto.Locked = True
    txtPreco_Produto.Locked = True
    txtQuant_Estoque.Locked = True
    txtCod_Fornecedor.Locked = True
    txtNome_Fornecedor.Locked = False
    
    
  
End Sub

Private Sub Desbloquear()

    txtDescricao_Produto.Locked = False
    txtPreco_Produto.Locked = False
    txtQuant_Estoque.Locked = False
    txtCod_Fornecedor.Locked = False
    
  
    
End Sub

Private Sub cmdNovo_Click()
    Desbloquear
    Limpar
    
    txtDescricao_Produto.BackColor = &HFFFFFF
    txtPreco_Produto.BackColor = &HFFFFFF
    txtQuant_Estoque.BackColor = &HFFFFFF
    txtCod_Fornecedor.BackColor = &HFFFFFF
    
    
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdImprimir.Visible = False
    cmdExcluir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
    cmdProximo.Visible = False
    cmdAnterior.Visible = False
    cmdUltimo.Visible = False
    cmdPrimeiro.Visible = False
    cmdPesquisa.Visible = False
    
    txtCod_Produto.Text = Format(Val(varProCod) + 1, "00000")
    txtDescricao_Produto.SetFocus
    Controle = True
    
End Sub

Private Sub cmdAlterar_Click()
    Desbloquear
    
    txtDescricao_Produto.BackColor = &HFFFFFF
    txtPreco_Produto.BackColor = &HFFFFFF
    txtQuant_Estoque.BackColor = &HFFFFFF
    txtCod_Fornecedor.BackColor = &HFFFFFF
    
    
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdImprimir.Visible = False
    cmdExcluir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
    cmdProximo.Visible = False
    cmdAnterior.Visible = False
    cmdUltimo.Visible = False
    cmdPrimeiro.Visible = False
    cmdPesquisa.Visible = False
 
    txtDescricao_Produto.SetFocus
    Controle = False
    
End Sub

Private Sub Limpar()
    txtCod_Produto.Text = ""
    txtDescricao_Produto.Text = ""
    txtPreco_Produto.Text = ""
    txtQuant_Estoque.Text = ""
    txtCod_Fornecedor.Text = ""
    txtNome_Fornecedor.Text = ""
End Sub

Private Sub cmdCancelar_Click()
    
    txtDescricao_Produto.BackColor = &HFCF0E7
    txtPreco_Produto.BackColor = &HFCF0E7
    txtQuant_Estoque.BackColor = &HFCF0E7
    txtCod_Fornecedor.BackColor = &HFCF0E7
    
    Bloquear
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdImprimir.Visible = True
    cmdExcluir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdProximo.Visible = True
    cmdAnterior.Visible = True
    cmdUltimo.Visible = True
    cmdPrimeiro.Visible = True
    cmdPesquisa.Visible = True

    ExibirDados
End Sub

Private Sub Form_Load()
    Bloquear
    cmdSalvar.Visible = False
    cmdCancelar.Visible = True
    AbrirDB
    ExibirDados
End Sub



Private Sub cmdPrimeiro_Click()
    If tblProduto.BOF = True And tblProduto.EOF = True Then
        MsgBox "O Registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Primeiro
    tblProduto.MoveFirst
    ExibirDados
    Exit Sub
Primeiro:
    If Err.Number = 3021 Then
        MsgBox "Você já está no primeiro registro!", vbInformation, "Aviso!"
        tblProduto.MoveFirst
        ExibirDados
        Exit Sub
    End If
    
End Sub


Private Sub cmdUltimo_Click()
    If tblProduto.BOF = True And tblProduto.EOF = True Then
        MsgBox "O Registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Ultimo
    tblProduto.MoveLast
    ExibirDados
    Exit Sub
Ultimo:
    If Err.Number = 3021 Then
        MsgBox "Você já está no último registro!", vbInformation, "Aviso!"
        tblProduto.MoveLast
        ExibirDados
        Exit Sub
    End If
End Sub

Private Sub cmdAnterior_Click()
    If tblProduto.BOF = True And tblProduto.EOF = True Then
        MsgBox "O Registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Anterior
    tblProduto.MovePrevious
    If tblProduto.BOF = False Then
        ExibirDados
    End If
    Exit Sub
Anterior:
    If Err.Number = 3021 Then
        MsgBox "Você já está no primeiro Registro!", vbInformation, "Aviso!"
        tblProduto.MoveFirst
        ExibirDados
        Exit Sub
    End If
End Sub

Private Sub cmdProximo_Click()
    If tblProduto.BOF = True And tblProduto.EOF = True Then
        MsgBox "O Registro está vazio", vbExclamation, "Aviso!"
        Exit Sub
    End If
    On Error GoTo Proximo
    tblProduto.MoveNext
    If tblProduto.EOF = False Then
        ExibirDados
    End If
    Exit Sub
Proximo:
    If Err.Number = 3021 Then
        MsgBox "Você já está no último registro!", vbInformation, "Aviso!"
        tblProduto.MoveLast
        ExibirDados
        Exit Sub
    End If
End Sub




Private Sub cmdExcluir_Click()
    Dim Excluir As String
    Excluir = MsgBox("Deseja excluir este registro?", vbQuestion + vbYesNo, "Produto")
    If Excluir <> vbYes Then
        Exit Sub
    Else
        If Not tblProduto.EOF = True And Not tblProduto.BOF = True Then
            ACelula.BeginTrans
            tblProduto.Delete
            tblProduto.MovePrevious
            ExibirDados
            ACelula.CommitTrans
         Else
            Excluir = MsgBox("Fim de arquivo encontrado!", vbInformation, "Produto")
         End If
    End If
End Sub


Private Sub AbrirDB()
    Set ACelula = DBEngine.Workspaces(0)
    On Error GoTo ErroAbrir
    ACelula.BeginTrans
    Set BCelula = ACelula.OpenDatabase(App.Path & "\Celula.mdb", False)
    Set tblProduto = BCelula.OpenRecordset("tblProduto", dbOpenDynaset)
    Set TbProduto = tblProduto.OpenRecordset()
    Set tblCodigos = BCelula.OpenRecordset("tblcodigos", dbOpenDynaset)
    Set TbCodigos = tblCodigos.OpenRecordset()
    varProCod = tblCodigos!procod
    
    ACelula.CommitTrans
    On Error GoTo 0
    Exit Sub
ErroAbrir:
    Dim Aviso As Integer
    Aviso = MsgBox("Erro ao acessar os Dados!", vbCritical, "Aviso!!!")
    If Aviso = vbOK Then
        ACelula.Rollback
        On Error GoTo 0
        Exit Sub
    End If
    Resume
End Sub

Private Sub cmdSalvar_Click()
    ACelula.BeginTrans ' inicia transação com o banco de dados
    If Controle = True Then
        tblProduto.AddNew
        tblProduto!Cod_Produto = txtCod_Produto.Text
        tblProduto!Descricao_Produto = txtDescricao_Produto.Text
        tblProduto!Preco_Produto = txtPreco_Produto.Text
        tblProduto!Quant_Estoque = txtQuant_Estoque.Text
        tblProduto!Cod_Fornecedor = txtCod_Fornecedor.Text
        tblProduto!nome_fornecedor = txtNome_Fornecedor.Text

        tblProduto.Update
        MsgBox " Inclusão com sucesso!", vbInformation, " Produto"
        varProCod = Val(varProCod) + 1
        
    Else
        tblProduto.Edit
        tblProduto!Cod_Produto = txtCod_Produto.Text
        tblProduto!Descricao_Produto = txtDescricao_Produto.Text
        tblProduto!Preco_Produto = txtPreco_Produto.Text
        tblProduto!Quant_Estoque = txtQuant_Estoque.Text
        tblProduto!Cod_Fornecedor = txtCod_Fornecedor.Text
        tblProduto!nome_fornecedor = txtNome_Fornecedor.Text

        tblProduto.Update
        MsgBox " Alterado com sucesso!", vbInformation, " Produto"
        
    End If
    
        ACelula.CommitTrans
        
    txtDescricao_Produto.BackColor = &HFCF0E7
    txtPreco_Produto.BackColor = &HFCF0E7
    txtQuant_Estoque.BackColor = &HFCF0E7
    txtCod_Fornecedor.BackColor = &HFCF0E7

        
    Bloquear
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    cmdImprimir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    cmdProximo.Visible = True
    cmdAnterior.Visible = True
    cmdUltimo.Visible = True
    cmdPrimeiro.Visible = True
    cmdPesquisa.Visible = True
    
    
    Controle = True ' Deixa habilitado pois será inserido novo registro
    
    
End Sub

Private Sub ExibirDados()
    Limpar
    On Error Resume Next
    txtCod_Produto.Text = tblProduto!Cod_Produto
    txtDescricao_Produto.Text = tblProduto!Descricao_Produto
    txtPreco_Produto.Text = tblProduto!Preco_Produto
    txtQuant_Estoque.Text = tblProduto!Quant_Estoque
    txtCod_Fornecedor.Text = tblProduto!Cod_Fornecedor
    txtNome_Fornecedor.Text = tblProduto!nome_fornecedor
    On Error GoTo 0
End Sub
