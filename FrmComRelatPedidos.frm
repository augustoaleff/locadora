VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComRelatPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Pedidos"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox CmbStatus 
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox TxtCodCliente 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame FrameConsultaPor 
      Caption         =   "Consulta Por"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton OptDataEmissao 
         Caption         =   "Data Emissao"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptDataDev 
         Caption         =   "Data Devolução"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   200
         Width           =   1695
      End
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
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
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2520
      TabIndex        =   13
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label LblStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód Cliente:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "FrmComRelatPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdLimparTela_Click()
    Call limpa_campos
End Sub

Private Sub Command1_Click()
    FrmComRelatPedidosBuscarNome.Show
End Sub

Private Sub Form_Load()

    Me.Top = 1500
    Me.Left = 2000

    Call carregar_combo

End Sub
Private Sub limpa_campos()

    OptDataEmissao.Value = True
    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    CmbStatus.Text = ""
    MskPeriodoDe.Mask = ""
    MskPeriodoDe.Text = ""
    MskPeriodoDe.Mask = "##/##/####"
    MskPeriodoAte.Mask = ""
    MskPeriodoAte.Text = ""
    MskPeriodoAte.Mask = "##/##/####"

    TxtCodCliente.SetFocus

End Sub
Private Sub carregar_combo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM STATUS_RELAT_PED order by descricao")

    Do Until reg.EOF = True

        CmbStatus.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub

Private Sub CmbStatus_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbStatus.Text <> Empty Then

        MskPeriodoDe.SetFocus

    End If


End Sub
Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) Or Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty) Then

        MskPeriodoAte.SetFocus

    End If


End Sub
Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 And (IsDate(MskPeriodoAte.Text) Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty) Then

        CmdImprimir.SetFocus

    End If


End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT NOME FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodCliente.Text))


        If reg.EOF = False Then

            LblCliente.Caption = reg.Fields("Nome")
            CmbStatus.SetFocus

        Else

            MsgBox "Cliente não encontrado"

        End If

        reg.Close

    ElseIf KeyAscii = 13 And TxtCodCliente.Text = Empty Then

        CmbStatus.SetFocus
        LblCliente.Caption = ""


    ElseIf KeyAscii = 13 Then

        MsgBox "Digite um Código de Cliente Válido"
        LblCliente.Caption = ""

    End If


End Sub
