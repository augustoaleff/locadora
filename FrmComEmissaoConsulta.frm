VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Pedidos"
   ClientHeight    =   7575
   ClientLeft      =   7740
   ClientTop       =   2175
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   13020
   Begin VB.CommandButton CmdComEmissaoConsultaImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   6000
      TabIndex        =   21
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton CmdComEmissaoConsultaPesquisaNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton CmdComEmissaoConsultaConsultar 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdComEmissaoConsultaLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   11040
      TabIndex        =   19
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame FrmComEmissaoConsultaDetalharPedido 
      Caption         =   "Detalhar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   6120
      Width           =   5415
      Begin VB.CommandButton CmdComEmissaoConsultaDetalhar 
         Caption         =   "Detalhar"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtComEmissaoConsultaNumeroPedido 
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
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label LblComEmissaoConsultaNumeroPedido 
         Caption         =   "Nª Pedido"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame FrmComEmissaoConsultaStatusAluguel 
      Caption         =   "Status Aluguel"
      Height          =   735
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   3975
      Begin VB.OptionButton OptComEmissaoConsultaAguardandoDevolucao 
         Caption         =   "Aguardando Devolução"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton OptComEmissaoConsultaDevolvidos 
         Caption         =   "Devolvidos"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptComEmissaoConsultaTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtComEmissaoConsultaPeriodoAte 
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox TxtComEmissaoConsultaPeriodoDe 
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
      Top             =   1200
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComEmissaoConsultaPedidos 
      Height          =   4095
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.TextBox TxtComEmissaoConsultaCodCliente 
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
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LblComEmissaoConsultaValorTotal 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   450
      Left            =   10560
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label LblComEmissaoConsultaTotal 
      Caption         =   "Total R$"
      Height          =   255
      Left            =   9720
      TabIndex        =   17
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblComEmissaoConsultaA 
      Caption         =   "à"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label LblComEmissaoConsultaPeriodo 
      Caption         =   "Período"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblComEmissaoConsultaCliente 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label LblComEmissaoConsultaCodCliente 
      Caption         =   "Cód . Cliente"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   945
   End
End
Attribute VB_Name = "FrmComEmissaoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()

End Sub

