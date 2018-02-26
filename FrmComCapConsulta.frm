VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCapConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Contas a Pagar"
   ClientHeight    =   7545
   ClientLeft      =   7920
   ClientTop       =   1170
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11490
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   1800
      Width           =   3375
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptAberto 
         Caption         =   "Aberto"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptPago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox CmbTipoPagto 
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
      ItemData        =   "FrmComCapConsulta.frx":0000
      Left            =   7560
      List            =   "FrmComCapConsulta.frx":0002
      TabIndex        =   18
      Top             =   2040
      Width           =   3735
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   375
      Left            =   1080
      TabIndex        =   16
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
   Begin VB.CommandButton CmdConsultarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   4920
      TabIndex        =   12
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   9120
      TabIndex        =   2
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Relatório"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexResultado 
      Height          =   3735
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.TextBox TxtCodForn 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame FrameConsultarPor 
      Caption         =   "Consultar Por"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.OptionButton OptDataLancamento 
         Caption         =   "Data Lançamento"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptDataPagamento 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   375
      Left            =   2640
      TabIndex        =   17
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
   Begin VB.Label LblTipoPagto 
      Caption         =   "Tipo Pagto"
      Height          =   255
      Left            =   7560
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LblValor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label LblValorTotal 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   8160
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label LblForn 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label LblPeriodoA 
      Caption         =   "à"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LblCodForn 
      Caption         =   "Cód Forn"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "FrmComCapConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FILTROS, CONSULTA, STATUS, TIPO, FORN, DATADE, DATAATE As String
Attribute CONSULTA.VB_VarUserMemId = 1073938432
Attribute STATUS.VB_VarUserMemId = 1073938432
Attribute TIPO.VB_VarUserMemId = 1073938432
Attribute FORN.VB_VarUserMemId = 1073938432
Attribute DATADE.VB_VarUserMemId = 1073938432
Attribute DATAATE.VB_VarUserMemId = 1073938432

Private Sub CmdConsultar_Click()
    Dim VTOTAL As Double

    Call WHERE


    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT * FROM C_A_P AS P FULL OUTER JOIN FORNECEDORES AS F ON F.CODFORN = P.codforn WHERE " & FILTROS & "")

    Call formata_flex

    Do Until reg.EOF = True

        If reg.Fields("Status") = "P" Then


            MSFlexResultado.AddItem (reg.Fields("Vencto") & vbTab & _
                                     reg.Fields("CodForn") & vbTab & _
                                     reg.Fields("RazaoSocial") & vbTab & _
                                     reg.Fields("NumDocto") & vbTab & _
                                     reg.Fields("Tipo") & vbTab & _
                                     Format(reg.Fields("Valor"), "#,##0.00") & vbTab & _
                                     reg.Fields("Status") & vbTab & _
                                     Format(reg.Fields("DataPagto"), "DD/MM/YYYY") & vbTab & _
                                     reg.Fields("Obs"))
        Else

            MSFlexResultado.AddItem (reg.Fields("Vencto") & vbTab & _
                                     reg.Fields("CodForn") & vbTab & _
                                     reg.Fields("RazaoSocial") & vbTab & _
                                     reg.Fields("NumDocto") & vbTab & _
                                     reg.Fields("Tipo") & vbTab & _
                                     Format(reg.Fields("Valor"), "#,##0.00") & vbTab & _
                                     reg.Fields("Status") & vbTab & _
                                     vbTab & _
                                     reg.Fields("Obs"))
        End If


        reg.MoveNext

    Loop

    reg.Close


    For contador = 1 To MSFlexResultado.Rows - 1

        VTOTAL = VTOTAL + CDbl(MSFlexResultado.TextMatrix(contador, 5))

    Next

    LblValor.Caption = Format(VTOTAL, "#,##0.00")

End Sub

Private Sub CmdConsultarNome_Click()

    FrmComCapConsultaBuscarNome.Show

End Sub


Private Sub Form_Load()

    Call carregar_combo_tipo

End Sub
Private Sub WHERE()

    CONSULTA = ""
    STATUS = ""
    TIPO = ""
    FILTROS = ""


    If OptVencimento.Value = True Then

        CONSULTA = "VENCTO BETWEEN "

    ElseIf OptDataPagamento.Value = True Then

        CONSULTA = "DATAPAGTO BETWEEN "

    ElseIf OptDataLancamento.Value = True Then

        CONSULTA = "DATALANCTO BETWEEN "

    End If


    If OptAberto.Value = True Then

        STATUS = " AND STATUS = 'A'"

    ElseIf OptPago.Value = True Then

        STATUS = " AND STATUS = 'P'"

    ElseIf OptTodos.Value = True Then

        STATUS = ""

    End If

    If CmbTipoPagto.Text = "Boleto" Then

        TIPO = " AND TIPO = 'BOL'"

    ElseIf CmbTipoPagto.Text = "Transferencia" Then

        TIPO = " AND TIPO = 'TRANSF'"

    ElseIf CmbTipoPagto.Text = "Deposito" Then

        TIPO = " AND TIPO = 'DEP'"

    ElseIf CmbTipoPagto.Text = "Cheque" Then

        TIPO = " AND TIPO = 'CHQ'"

    ElseIf CmbTipoPagto.Text = "Cartao Credito" Then

        TIPO = " AND TIPO = 'CC'"

    ElseIf CmbTipoPagto.Text = "Cartao Debito" Then

        TIPO = " AND TIPO = 'CD'"

    ElseIf CmbTipoPagto.Text = "Dinheiro" Then

        TIPO = " AND TIPO = 'DH'"

    ElseIf CmbTipoPagto.Text = "*Todos" Then

        TIPO = ""

    Else

        TIPO = ""

    End If


    If TxtCodForn.Text <> Empty Then

        FORN = " AND P.CODFORN = " & Trim(TxtCodForn.Text)

    Else

        FORN = ""

    End If

    If Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty Then

        DATADE = Format("01/01/1900", "YYYYMMDD")

    Else

        DATADE = Format(MskPeriodoDe.Text, "YYYYMMDD")

    End If

    If Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty Then

        DATAATE = Format("31/12/2199", "YYYYMMDD")

    Else

        DATAATE = Format(MskPeriodoAte.Text, "YYYYMMDD")

    End If

    FILTROS = CONSULTA + "'" + DATADE + "' AND '" + DATAATE + "'" + FORN + STATUS + TIPO


End Sub
Private Sub carregar_combo_tipo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM TIPOS_PAGTOS_CAP_CST order by descricao")

    Do Until reg.EOF = True

        CmbTipoPagto.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub

Public Sub TxtCodForn_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodForn.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        reg.Open ("SELECT RazaoSocial FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodForn.Text) & "")

        If reg.EOF = False Then

            LblForn.Caption = reg.Fields("RazaoSocial")
            MskPeriodoDe.SetFocus

        Else

            MsgBox "Fornecedor não encontrado"

        End If

    ElseIf KeyAscii = 13 And TxtCodForn.Text = Empty Then

        MskPeriodoDe.SetFocus
        LblForn.Caption = ""


    ElseIf KeyAscii = 13 Then

        MsgBox "Digite um código de fornecedor válido"
        LblForn.Caption = ""

    End If


End Sub


Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) <> Empty Or Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty) Then

        MskPeriodoAte.SetFocus

    End If


End Sub

Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 And (IsDate(MskPeriodoAte.Text) <> Empty Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty) Then

        CmdConsultar.SetFocus

    End If


End Sub

Private Sub formata_flex()

    MSFlexResultado.Clear
    MSFlexResultado.Cols = 9
    MSFlexResultado.Rows = 1

    MSFlexResultado.Col = 0
    MSFlexResultado.Text = "Vencto"
    MSFlexResultado.ColWidth(0) = 1000

    MSFlexResultado.Col = 1
    MSFlexResultado.Text = "Cod. Forn"
    MSFlexResultado.ColWidth(1) = 1000

    MSFlexResultado.Col = 2
    MSFlexResultado.Text = "Fornecedor"
    MSFlexResultado.ColWidth(2) = 3000

    MSFlexResultado.Col = 3
    MSFlexResultado.Text = "Num. Docto"
    MSFlexResultado.ColWidth(3) = 1500

    MSFlexResultado.Col = 4
    MSFlexResultado.Text = "Tipo"
    MSFlexResultado.ColWidth(4) = 500

    MSFlexResultado.Col = 5
    MSFlexResultado.Text = "Valor"
    MSFlexResultado.ColWidth(5) = 1000

    MSFlexResultado.Col = 6
    MSFlexResultado.Text = "Status"
    MSFlexResultado.ColWidth(6) = 500

    MSFlexResultado.Col = 7
    MSFlexResultado.Text = "Data. Pagto"
    MSFlexResultado.ColWidth(7) = 1000

    MSFlexResultado.Col = 8
    MSFlexResultado.Text = "OBS"
    MSFlexResultado.ColWidth(8) = 3000


End Sub

