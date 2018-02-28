VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComRelatCAR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relátório de Contas a Receber"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.ComboBox CmbStatus 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
   Begin VB.TextBox TxtCodCliente 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame FrameConsultarPor 
      Caption         =   "Consultar por"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton OptDataEmissao 
         Caption         =   "Data Emissão"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptDataPagto 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   330
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      Left            =   2400
      TabIndex        =   16
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2400
      TabIndex        =   15
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód.Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label LblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label LblTipo 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
End
Attribute VB_Name = "FrmComRelatCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub carregar_combo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM TIPOS_PAGTOS order by descricao")

    Do Until reg.EOF = True

        CmbTipo.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

    reg.Open ("SELECT Descricao FROM STATUS_RELAT_CAR order by descricao")

    Do Until reg.EOF = True

        CmbStatus.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub
Private Sub limpa_campos()

    OptVencimento.Value = True

    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    MskPeriodoDe.Mask = ""
    MskPeriodoDe.Text = ""
    MskPeriodoDe.Mask = "##/##/####"
    MskPeriodoAte.Mask = ""
    MskPeriodoAte.Text = ""
    MskPeriodoAte.Mask = "##/##/####"
    CmbTipo.Text = ""
    CmbStatus.Text = ""

    TxtCodCliente.SetFocus


End Sub


Private Sub CmdBuscarNome_Click()

    FrmComRelatCARBuscarNome.Show

End Sub

Private Sub CmdLimparTela_Click()

    Call limpa_campos

End Sub

Private Sub Form_Load()

    Call carregar_combo

    Me.Top = 1000
    Me.Left = 2000

End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT NOME FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodCliente.Text) & "")

        If reg.EOF = False Then

            LblCliente.Caption = reg.Fields("Nome")

            MskPeriodoDe.SetFocus

        Else

            MsgBox "Cliente não encontrado"

        End If

        reg.Close

    ElseIf KeyAscii = 13 And TxtCodCliente.Text = Empty Then

        LblCliente.Caption = ""
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

        CmbTipo.SetFocus

    End If


End Sub


Private Sub CmbTipo_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then

        CmbStatus.SetFocus

    End If


End Sub

Private Sub CmbStatus_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then

        CmdImprimir.SetFocus

    End If


End Sub

