VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoDevolucao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolução de Filme"
   ClientHeight    =   8400
   ClientLeft      =   3750
   ClientTop       =   2865
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11430
   Begin VB.TextBox TxtComEmissaoDevolucaoQuant 
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
      Left            =   8160
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCodProduto 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoBaixarItem 
      Caption         =   "Baixar Item"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCodVendedor 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoPesquisa 
      Caption         =   "Pesquisar Pedido por Cliente"
      Height          =   495
      Left            =   4800
      TabIndex        =   47
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   375
      Left            =   480
      TabIndex        =   46
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoFecharDevolucao 
      Caption         =   "Fechar Devolução"
      Height          =   615
      Left            =   9120
      TabIndex        =   17
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoDiasAtraso 
      Enabled         =   0   'False
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
      Left            =   10320
      TabIndex        =   45
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoDataLimiteDevolucao 
      Enabled         =   0   'False
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
      Left            =   5760
      TabIndex        =   44
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoDataEntrega 
      Enabled         =   0   'False
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
      TabIndex        =   43
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCodCliente 
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   42
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCodCupomDesconto 
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCheque 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCartaoCredito 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoCartaoDebito 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoValorTotal 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   33
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame FrameComEmissaoDevolucaoFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2775
      Left            =   360
      TabIndex        =   23
      Top             =   4560
      Width           =   6255
      Begin VB.TextBox TxtComEmissaoDevolucaoDinheiro 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtComEmissaoDevolucaoParcelasCC 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox CmbComEmissaoAluguelBandeiraCC 
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox CmbComEmissaoAluguelBandeiraCD 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CmdComEmissaoAluguelValidarCupom 
         Caption         =   "Validar Cupom"
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtComEmissaoDevolucaoQuantCheque 
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label LblComEmissaoDevolucaoDinheiro 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblComEmissaoDevolucaoCartaoDebito 
         Caption         =   "Cartão de Débito"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblComEmissaoDevolucaoCartaoCredito 
         Caption         =   "Cartão de Crédito"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblComEmissaoDevolucaoCheque 
         Caption         =   "Cheque"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblComEmissaoDevolucaoCodCupomDesconto 
         Caption         =   "Cód. Cupom de Desconto"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblComEmissaoDevolucaoParcelasCC 
         Caption         =   "Parcelas"
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblComEmissaoDevolucaoBandeiraCD 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblComEmissaoDevolucaoBandeiraCC 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblComEmissaoDevolucaoQuantCheque 
         Caption         =   "Quant."
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoValorPago 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9360
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoMulta 
      Enabled         =   0   'False
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
      Left            =   9360
      TabIndex        =   21
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoDiferenca 
      Enabled         =   0   'False
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
      Left            =   9360
      TabIndex        =   20
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoDataDevolucao 
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
      Left            =   9240
      TabIndex        =   3
      Text            =   "__/__/____"
      Top             =   240
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComEmissaoDevolucaoItens 
      Height          =   1815
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoNumeroPedido 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label LblComEmissaoDevolucaoCliente 
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
      Left            =   2880
      TabIndex        =   53
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label LblComEmissaoDevolucaoProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3240
      TabIndex        =   52
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label LblComEmissaoDevolucaoQuant 
      Caption         =   "Quant."
      Height          =   255
      Left            =   7440
      TabIndex        =   51
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblComEmissaoDevolucaoCodProduto 
      Caption         =   "Cód Produto"
      Height          =   255
      Left            =   960
      TabIndex        =   50
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LblComEmissaoDevolucaoVendedor 
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
      Left            =   5520
      TabIndex        =   49
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label LblComEmissaoDevolucaoCodVendedor 
      Caption         =   "Cód.Vendedor"
      Height          =   255
      Left            =   3600
      TabIndex        =   48
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label LblComEmissaoDevolucaoDiasAtraso 
      Caption         =   "Dias de Atraso"
      Height          =   255
      Left            =   9000
      TabIndex        =   41
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblComEmissaoDevolucaoDataLimiteDevolucao 
      Caption         =   "Data Limite Devolução"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LblComEmissaoDevolucaoDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label LblComEmissaoDevolucaoCodCliente 
      Caption         =   "Cód. Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblComEmissaoDevolucaoValorTotal 
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
      Left            =   7200
      TabIndex        =   37
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label LblComEmissaoDevolucaoValorPago 
      Caption         =   "Valor Pago R$"
      Height          =   375
      Left            =   7920
      TabIndex        =   36
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label LblComEmissaoDevolucaoDiferenca 
      Caption         =   "Diferença R$"
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
      Left            =   7560
      TabIndex        =   35
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label LblComEmissaoDevolucaoMulta 
      Caption         =   "Multa/Juros R$"
      Height          =   375
      Left            =   7920
      TabIndex        =   34
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label LblComEmissaoDevolucaoDataDevolucao 
      Caption         =   "Data Devolução"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label LblComEmissaoDevolucaoPedido 
      Caption         =   "Nº Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmComEmissaoDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub LblComAluguelDesconto_Click()

End Sub
Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub LblComAluguelCodVendedor_Click()

End Sub

Private Sub LblComAluguelDinheiro_Click()

End Sub

Private Sub LblComAluguelParcelasCC_Click()

End Sub

Private Sub LblComAluguelPgtoMinimo_Click()

End Sub

Private Sub Text7_Change()

End Sub

Private Sub LblComDevolucaoVendedor_Click()

End Sub

