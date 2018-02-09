VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.CommandButton CmdComEmissaoDevolucaoBuscarPedidoBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoBuscarPedidoLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   6840
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComEmissaoDevolucaoBuscarPedidoPedidos 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoBuscarPedidoAte 
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoBuscarPedidoPeriodoDe 
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
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdComEmissaoDevolucaoBuscarPedidoBuscarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtComEmissaoDevolucaoBuscarPedidoCodCliente 
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
   Begin VB.Label LblComEmissaoDevolucaoBuscarPedidoA 
      Caption         =   "a"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label LblComEmissaoDevolucaoBuscarPedidoCliente 
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
      TabIndex        =   9
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label LblComEmissaoDevolucaoBuscarPedidoPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblComEmissaoDevolucaoBuscarPedidoCodCliente 
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
Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

