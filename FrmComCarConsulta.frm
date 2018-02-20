VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCarConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Contas a Receber"
   ClientHeight    =   7500
   ClientLeft      =   930
   ClientTop       =   6240
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11640
   Begin VB.TextBox TxtComCarConsultaNumeroPedido 
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
      Left            =   9720
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame FrameComCarConsultaConsultarPor 
      Caption         =   "Consultar Por"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      Begin VB.OptionButton OptComCarConsultaVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptComCarConsultaDataPagamento 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptComCarConsultaDataLancamento 
         Caption         =   "Data Lançamento"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox TxtComCarConsultaCodForn 
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
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtComCarConsultaPeriodoDe 
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtComCarConsultaPeriodoAte 
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
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtComCarConsultaStatus 
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
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox TxtComCarConsultaTipoPagto 
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
      Left            =   6960
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton CmdComCarConsutaImprimir 
      Caption         =   "Imprimir Relatório"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton CmdComCarConsutaConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   9360
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCarConsutaLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCarConsultaConsultarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComCarConsultaResultado 
      Height          =   3735
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.Label LblComCarConsultaNumeroPedido 
      Caption         =   "Pedido nº"
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label LblComCarConsultaCodCliente 
      Caption         =   "Cód Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblComCarConsultaPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblComCarConsultaPeriodoA 
      Caption         =   "à"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label LblComCarConsultaForn 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2400
      TabIndex        =   19
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label LblComCarConsultaStatus 
      Caption         =   "Status"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblComCarConsultaTipoPagto 
      Caption         =   "Tipo Pagto"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LblComCarConsultaValorTotal 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   8400
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LblComCarConsultaValor 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   9480
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "FrmComCarConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LblComCapConsultaCodCliente_Click()

End Sub
