VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoConsultaPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Pedido"
   ClientHeight    =   9165
   ClientLeft      =   1695
   ClientTop       =   735
   ClientWidth     =   10335
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   10335
   Begin VB.CommandButton CmdComEmissaoConsultaPedidoImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3720
      TabIndex        =   44
      Top             =   8400
      Width           =   2295
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoDataDevolucao 
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
      Left            =   6480
      TabIndex        =   42
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CmdComEmissaoConsultaPedidoLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   8280
      Width           =   2775
   End
   Begin VB.CommandButton CmdComEmissaoConsultaPedidoCancelarPedido 
      Caption         =   "Cancelar Pedido"
      Height          =   375
      Left            =   240
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Frame FrameComEmissaoConsultaPedidoFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2415
      Left            =   1680
      TabIndex        =   28
      Top             =   5640
      Width           =   6975
      Begin VB.TextBox TxtComEmissaoConsultaPedidoTotalDesconto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   1920
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexComEmissaoConsultaPedidoCupons 
         Height          =   1695
         Left            =   3840
         TabIndex        =   38
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2990
         _Version        =   393216
      End
      Begin VB.TextBox TxtComEmissaoConsultaPedidoTotalCheque 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TxtComEmissaoConsultaPedidoTotalCC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtComEmissaoConsultaPedidoTotalCD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtComEmissaoConsultaPedidoTotalDinheiro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label LblComEmissaoConsultaPedidoTotalDesconto 
         Caption         =   "Total Desconto R$"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label LblComEmissaoConsultaPedidoCuponsUtilizados 
         Caption         =   "Cupons Utilizados"
         Height          =   375
         Left            =   4560
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblComEmissaoConsultaPedidoTotalCheque 
         Caption         =   "Total Cheque R$"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label LblComEmissaoConsultaPedidoTotalCC 
         Caption         =   "Total Cartão Crédito R$"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label LblComEmissaoConsultaPedidoTotalCD 
         Caption         =   "Total Cartão Débito R$"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label LblComEmissaoConsultaPedidoTotalDinheiro 
         Caption         =   "Total Dinheiro R$"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoStatusPedido 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7560
      TabIndex        =   27
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoValorAReceber 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   25
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoValorPago 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   24
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoOBS 
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
      Left            =   960
      TabIndex        =   22
      Top             =   1920
      Width           =   8895
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoValorTotal 
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
      Left            =   2640
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComEmissaoConsultaPedidoItens 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoVendedor 
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
      Left            =   5160
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoCodVendedor 
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
      Left            =   4200
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoDiasAtrasos 
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
      Left            =   9240
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoDataLimiteDevolucao 
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
      Left            =   3600
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoDataEntrega 
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
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoDataEmissao 
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
      Left            =   8760
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoCodCliente 
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
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoCliente 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox TxtComEmissaoConsultaPedidoNumeroPedido 
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
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label LblComEmissaoConsultaPedidoDataDevolucao 
      Alignment       =   2  'Center
      Caption         =   "Data Devolução"
      Height          =   375
      Left            =   5400
      TabIndex        =   41
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblComEmissaoConsultaPedidosStatusPedido 
      Caption         =   "Status Pedido"
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
      Left            =   8160
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblComEmissaoConsultaPedidoOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label LblComEmissaoConsultaPedidoValorAReceber 
      Caption         =   "Valor À Receber R$"
      Height          =   255
      Left            =   6480
      TabIndex        =   21
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label LblComEmissaoConsultaPedidoValorPago 
      Caption         =   "Valor Pago R$"
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label LblComEmissaoConsultaPedidoValorTotal 
      Caption         =   "Valor Total Pedido R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label LblComEmissaoConsultaPedidoVendedor 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LblComEmissaoConsultaPedidoDiasAtraso 
      Caption         =   "Dias Atraso"
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblComEmissaoConsultaPedidoDataLimiteDevolucao 
      Caption         =   "Data Limite Devolução"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblComEmissaoConsultaPedidoDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblComEmissaoConsultaPedidoDataEmissao 
      Caption         =   "Data Emissão"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblComEmissaoConsultaPedidoCodCliente 
      Caption         =   "Cód Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblComEmissaoConsultaPedidoNumeroPedido 
      Caption         =   "Pedido nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmComEmissaoConsultaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub

