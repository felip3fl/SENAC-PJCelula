VERSION 5.00
Begin VB.Form frmDescricao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descrição do Funcionario - CELL SOFT"
   ClientHeight    =   6600
   ClientLeft      =   9735
   ClientTop       =   2205
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DescricaoFuncionario.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   6255
   Begin VB.CommandButton cmdSair 
      Height          =   1215
      Left            =   4680
      Picture         =   "DescricaoFuncionario.frx":21CB6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sair da descrição  do funcionario"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   1215
      Left            =   3120
      Picture         =   "DescricaoFuncionario.frx":27535
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Excluir/Limpar o campo descrição do Funcionario"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   1215
      Left            =   1560
      Picture         =   "DescricaoFuncionario.frx":2BE83
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir a descrição do Funcionario"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   1215
      Left            =   3120
      Picture         =   "DescricaoFuncionario.frx":30DCE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Height          =   1215
      Left            =   1560
      Picture         =   "DescricaoFuncionario.frx":36A6E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salvar essa descrição de Funcionario"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   1215
      Left            =   0
      Picture         =   "DescricaoFuncionario.frx":3B38A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Editar o descrição do Funcionario"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtDescricao 
      BackColor       =   &H00FCF0E7&
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Digite aqui a descrição do funcionario"
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "- Conportamento"
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
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "- Deficiencia Fisica"
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
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "- Alegias"
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
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Doenças"
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
      TabIndex        =   10
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Doenças"
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
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Insira as descrições do funcionario como informações não presente no formulario, infomações como:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do Funcionario"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmDescricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    
    cmdEditar.Visible = True
    cmdImprimir.Visible = True
    cmdExcluir.Visible = True
    cmdSalvar.Visible = False
    cmdCancelar.Visible = False
    
    
    
    txtDescricao.BackColor = &HFCF0E7
    
End Sub

Private Sub cmdEditar_Click()

    

    txtDescricao.BackColor = &HFFFFFF
    
    cmdEditar.Visible = False
    cmdExcluir.Visible = False
    cmdImprimir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True
    txtDescricao.SetFocus
    Controle = False
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
