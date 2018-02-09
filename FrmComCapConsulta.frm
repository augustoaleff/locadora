VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.CommandButton CmdComCapConsultaConsultarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCapConsutaLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   4920
      TabIndex        =   18
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCapConsutaConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCapConsutaImprimir 
      Caption         =   "Imprimir Relatório"
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox TxtComCapConsultaTipoPagto 
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
      Left            =   7560
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TxtComCapConsultaStatus 
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
      Left            =   6000
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComCapConsultaResultado 
      Height          =   3735
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.TextBox TxtComCapConsultaPeriodoAte 
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
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TxtComCapConsultaPeriodoDe 
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
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TxtComCapConsultaCodForn 
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
   Begin VB.Frame FrameComCapConsultaConsultarPor 
      Caption         =   "Consultar Por"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.OptionButton OptComCapConsultaDataLancamento 
         Caption         =   "Data Lançamento"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptComCapConsultaDataPagamento 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptComCapConsultaVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label LblComCapConsultaValor 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   9240
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label LblComCapConsultaValorTotal 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label LblComCapConsultaTipoPagto 
      Caption         =   "Tipo Pagto"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label LblComCapConsultaStatus 
      Caption         =   "Status"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LblComCapConsultaForn 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2400
      TabIndex        =   13
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label LblComCapConsultaPeriodoA 
      Caption         =   "à"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label LblComCapConsultaPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LblComCapConsultaCodForn 
      Caption         =   "Cód Forn"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "FrmComCapConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

