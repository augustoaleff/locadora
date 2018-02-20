VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoDevolucaoBuscarPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar Pedido por Cliente"
   ClientHeight    =   7560
   ClientLeft      =   2400
   ClientTop       =   3750
   ClientWidth     =   7545
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7545
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   330
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPedidos 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1575
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
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   330
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "a"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   135
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
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód. Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FrmComEmissaoDevolucaoBuscarPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdBuscar_Click()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT NumPed,DataEntrega,Status,ValorT FROM PEDIDOS WHERE CODCLI= " & TxtCodCliente.Text & " and DataEntrega between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' AND '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "'")

    Call formata_flex

    Do Until reg.EOF = True

        MSFlexPedidos.AddItem (reg.Fields("NumPed") & vbTab & _
                               reg.Fields("DataEntrega") & vbTab & _
                               Format(reg.Fields("ValorT"), "#,##0.00") & vbTab & _
                               reg.Fields("Status"))

        reg.MoveNext

    Loop

    reg.Close

End Sub
Private Sub MSFlexPedidos_KeyPress(KeyAscii As Integer)

    Dim CODIGO As Long

    If KeyAscii = 13 Then

        MSFlexPedidos.Col = 0
        CODIGO = Trim(MSFlexPedidos.Text)

        FrmComEmissaoDevolucao.TxtNumPedido.Text = CODIGO
        FrmComEmissaoDevolucao.TxtNumPedido_KeyPress (13)
        Unload Me

    End If

End Sub


Private Sub CmdBuscarNome_Click()
    FrmComEmissaoDevolucaoBuscarPedidoBuscarNome.Show
End Sub

Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskPeriodoDe.Text) <> Empty Then

        MskPeriodoAte.SetFocus

    End If

End Sub
Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskPeriodoAte.Text) <> Empty Then

        CmdBuscar.SetFocus

    End If

End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT NOME FROM CLIENTES WHERE CODCLI = " & TxtCodCliente.Text & "")

        If reg.EOF = False Then

            LblCliente.Caption = reg.Fields("Nome")
            MskPeriodoDe.SetFocus

        Else

            MsgBox "Cliente Não Encontrado", vbInformation, Aviso

        End If

        reg.Close

    End If

End Sub
Private Sub formata_flex()

    MSFlexPedidos.Clear
    MSFlexPedidos.Cols = 4
    MSFlexPedidos.Rows = 1

    MSFlexPedidos.Col = 0
    MSFlexPedidos.Text = "Nº Pedido"
    MSFlexPedidos.ColWidth(0) = 1000

    MSFlexPedidos.Col = 1
    MSFlexPedidos.Text = "Data"
    MSFlexPedidos.ColWidth(1) = 1000

    MSFlexPedidos.Col = 2
    MSFlexPedidos.Text = "Valor total"
    MSFlexPedidos.ColWidth(2) = 1000

    MSFlexPedidos.Col = 3
    MSFlexPedidos.Text = "Status"
    MSFlexPedidos.ColWidth(3) = 1000


End Sub
