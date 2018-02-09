VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoAluguel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emiss�o de Aluguel de Filme"
   ClientHeight    =   9225
   ClientLeft      =   225
   ClientTop       =   2505
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11085
   Begin VB.CommandButton CmdGerarNumeroPedido 
      Caption         =   "Gerar n� Pedido"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar por Nome do Cliente"
      Height          =   495
      Left            =   4440
      TabIndex        =   55
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox TxtOBS 
      Height          =   330
      Left            =   840
      TabIndex        =   7
      Top             =   1920
      Width           =   10095
   End
   Begin VB.CommandButton CmdRemoverItem 
      Caption         =   "Remover Item"
      Height          =   375
      Left            =   9720
      TabIndex        =   52
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TxtDesconto 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   51
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox TxtDiferenca 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   49
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox TxtValorRecebido 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   48
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox TxtPagtoMinimo 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9000
      TabIndex        =   45
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox TxtCupomDesconto 
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TxtCheque 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoCredito 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoDebito 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame FrameFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2775
      Left            =   240
      TabIndex        =   35
      Top             =   5400
      Width           =   6255
      Begin VB.TextBox TxtQuantCheque 
         Height          =   285
         Left            =   3480
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton CmdValidarCupom 
         Caption         =   "Validar Cupom"
         Height          =   495
         Left            =   3480
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox CmbBandeiraCD 
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CmbBandeiraCC 
         Height          =   315
         Left            =   3480
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtParcelasCC 
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TxtDinheiro 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LblQuantCheque 
         Caption         =   "Quant."
         Height          =   255
         Left            =   2760
         TabIndex        =   53
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label LblBandeiraCC 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblBandeiraCD 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblParcelasCC 
         Caption         =   "Parcelas"
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblCodCupomDesconto 
         Caption         =   "C�d. Cupom de Desconto"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblCheque 
         Caption         =   "Cheque"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblCartaoCredito 
         Caption         =   "Cart�o de Cr�dito"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblCartaoDebito 
         Caption         =   "Cart�o de D�bito"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblDinheiro 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdCancelarPedido 
      Caption         =   "Cancelar Pedido"
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox TxtValorTotal 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9000
      TabIndex        =   32
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox PctFoto 
      Height          =   1815
      Left            =   9120
      ScaleHeight     =   1755
      ScaleWidth      =   1635
      TabIndex        =   31
      Top             =   3240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexItens 
      Height          =   1935
      Left            =   360
      TabIndex        =   30
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.TextBox TxtQuant 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdInserir 
      Caption         =   "Inserir"
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtCodProduto 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox TxtDataLimiteDevolucao 
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Text            =   "__/__/____"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TxtDataEntrega 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "__/__/____"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton CmdEmitirPedido 
      Caption         =   "Emitir Pedido"
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox TxtCodVendedor 
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtCodCliente 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox TxtNumPedido 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LblVendedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8400
      TabIndex        =   58
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label LblQuantDiasAlugados 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10080
      TabIndex        =   57
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblDiasAlugados 
      Caption         =   "Dias Alugados"
      Height          =   255
      Left            =   8880
      TabIndex        =   56
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label LblProduto 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label LblDesconto 
      Caption         =   "Desconto R$"
      Height          =   255
      Left            =   7680
      TabIndex        =   50
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label LblValorRecebido 
      Caption         =   "Valor Recebido R$"
      Height          =   375
      Left            =   7320
      TabIndex        =   47
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label LblDiferenca 
      Caption         =   "Diferen�a R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   46
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label LblPgtoMinimo 
      Caption         =   "Pagamento M�nimo R$"
      Height          =   375
      Left            =   6960
      TabIndex        =   39
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label LblValorTotal 
      Caption         =   "Valor Total R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   33
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label LblQuantProduto 
      Caption         =   "Quant."
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LblCodProduto 
      Caption         =   "C�d. Produto"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblDataLimiteDevolucao 
      Caption         =   "Data Limite Devolu��o"
      Height          =   375
      Left            =   4680
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label LblCodVendedor 
      Caption         =   "C�d.Vendedor"
      Height          =   255
      Left            =   6240
      TabIndex        =   25
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "C�d. Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LblNumPedido 
      Caption         =   "Pedido n�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmComEmissaoAluguel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGerarNumeroPedido_Click()

If TxtNumPedido.Text = Empty Then
 
 TxtNumPedido.Enabled = True
 
 Dim QUERY As String
 
 Set CN1 = New ADODB.Connection
 CN1.Open STR_DSN
 Set reg = New ADODB.Recordset
 reg.ActiveConnection = CN1
 
 
 CN1.Execute ("begin transaction")
 QUERY = "select UltNumPedido from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltNumPedido = UltNumPedido+7;COMMIT"
 reg.Open (QUERY)
 
 TxtNumPedido.Text = reg.Fields("UltNumPedido")
 TxtNumPedido.Enabled = False
 TxtCodVendedor.SetFocus
 
 reg.Close
 
 Else
 
  MsgBox "Emita ou Cancele o Pedido Atual para criar um Novo!", vbExclamation, "Aviso"
 
 End If
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Top = 100
End Sub
Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

 Set CN1 = New ADODB.Connection
 CN1.Open STR_DSN
 Set reg = New ADODB.Recordset
 reg.ActiveConnection = CN1
 
 reg.Open ("SELECT Nome FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodCliente.Text) & "")
 
 If reg.EOF = False Then
 
 LblCliente.Caption = reg.Fields("Nome")

 TxtDataEntrega.SetFocus
 
 Else
 
 MsgBox "Cliente N�o Encontrado", vbExclamation, "Aviso"
 TxtCodCliente.SetFocus
 
 End If
    
 reg.Close
 

End If

End Sub

Private Sub TxtCodVendedor_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsNumeric(TxtCodVendedor.Text) <> Empty Then

 Set CN1 = New ADODB.Connection
 CN1.Open STR_DSN
 Set reg = New ADODB.Recordset
 reg.ActiveConnection = CN1
 
 reg.Open ("SELECT Nome FROM FUNCIONARIOS WHERE CODFUNC = " & Trim(TxtCodVendedor.Text) & " AND CARGO LIKE '%VEND%'")
 
 If reg.EOF = False Then
 
 LblVendedor.Caption = reg.Fields("Nome")

 TxtCodCliente.SetFocus
 
 Else
 
 MsgBox "Vendedor N�o Existe ou n�o � Vendedor", vbExclamation, "Aviso"
 TxtCodVendedor.SetFocus
 
 End If
    
 reg.Close
 

End If

End Sub
