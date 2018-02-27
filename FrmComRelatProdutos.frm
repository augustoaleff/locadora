VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComRelatProdutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   4440
   ClientLeft      =   16170
   ClientTop       =   2535
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameConsultar 
      Caption         =   "Consultar:"
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton OptProdutosEmFalta 
         Caption         =   "Produtos Em Falta"
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton OptProdutosEmEstoque 
         Caption         =   "Produtos em Estoque"
         Height          =   435
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox TxtLocalizacao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox CmbCategoria 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox CmbTipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox TxtDescricao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label LblDataCadastro 
      Caption         =   "Data Cadastro"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LblDescricao 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label LblLocalizacao 
      Caption         =   "Localização"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label LblCategoria 
      Caption         =   "Categoria"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label LblTipo 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "FrmComRelatProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub carregar_combo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM TIPOS_PROD order by descricao")

    Do Until reg.EOF = True

        CmbTipo.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

    reg.Open ("SELECT Descricao FROM CATEGS_PROD order by descricao")

    Do Until reg.EOF = True

        CmbCategoria.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub

Private Sub limpa_campos()

    OptTodos.Value = True
    TxtDescricao.Text = ""
    CmbTipo.Text = ""
    CmbCategoria.Text = ""
    TxtLocalizacao.Text = ""
    MskPeriodoDe.Mask = ""
    MskPeriodoDe.Text = ""
    MskPeriodoDe.Mask = "##/##/####"
    MskPeriodoAte.Mask = ""
    MskPeriodoAte.Text = ""
    MskPeriodoAte.Mask = "##/##/####"
    
    TxtDescricao.SetFocus
    
 
End Sub

Private Sub CmdLimparTela_Click()
    
    Call limpa_campos
    
End Sub

Private Sub Form_Load()

    Call carregar_combo

End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        CmbTipo.SetFocus

    End If


End Sub

Private Sub CmbTipo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        CmbCategoria.SetFocus

    End If


End Sub
Private Sub CmbCategoria_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        TxtLocalizacao.SetFocus

    End If


End Sub
Private Sub TxtLocalizacao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        MskPeriodoDe.SetFocus

    End If


End Sub
Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) <> Empty Or Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty) Then

        MskPeriodoAte.SetFocus

    End If


End Sub

Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoAte.Text) <> Empty Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty) Then

        CmdImprimir.SetFocus

    End If

End Sub

