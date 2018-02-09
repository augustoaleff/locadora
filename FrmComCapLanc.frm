VERSION 5.00
Begin VB.Form FrmComCapLanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento Contas a Pagar"
   ClientHeight    =   5415
   ClientLeft      =   1320
   ClientTop       =   2565
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8400
   Begin VB.CommandButton CmdComCapLancBuscaNumeroDocumento 
      Caption         =   "Buscar por Nº Documento"
      Height          =   375
      Left            =   2400
      TabIndex        =   33
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtComCapLancSeq 
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
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CmdComCapLancBuscaNome 
      Caption         =   "Buscar Por Nome"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtComCapLancLocalPagto 
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
      Left            =   1560
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton CmdComCapLancLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton CmdComCapLancGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox TxtComCapLancDataLancamento 
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
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox TxtComCapLancOBS 
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
      TabIndex        =   10
      Top             =   3960
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   5280
      TabIndex        =   22
      Top             =   1560
      Width           =   2895
      Begin VB.TextBox TxtComCapLancValorTotal 
         Height          =   330
         Left            =   1320
         TabIndex        =   27
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtComCapLancJuros 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtComCapLancDesconto 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtComCapLancValor 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblComCapLancValorTotal 
         Caption         =   "Valor Total R$"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label LblComCapLancJuros 
         Caption         =   "Juros/Multa R$"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblComCapLancDesconto 
         Caption         =   "Desconto R$"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LblComCapLancValor 
         Caption         =   "Valor R$"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox TxtComCapLancDataPagto 
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
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox TxtComCapLancStatus 
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
      Left            =   1560
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox TxtComCapLancNumeroDocto 
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
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox CmbComCapLancTipoPagto 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox TxtComCapLancVencimento 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox TxtComCapLancForn 
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
      Left            =   3120
      TabIndex        =   15
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox TxtComCapLancCodForn 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label LblComCapLancSeq 
      Caption         =   "Seq."
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label LblComCapLancLocalPagto 
      Caption         =   "Local Pagto"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label LblComCapLancDataLancamento 
      Caption         =   "Data Lançamento"
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label LblComCapLancOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label LblComCapLancTipoPagto 
      Caption         =   "Tipo Pagamento"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label LblComCapLancNumeroDocto 
      Caption         =   "Nº Documento"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label LblComCapLancStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label LblComCapLancDataPagto 
      Caption         =   "Data Pagto"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LblComCapLancVencimento 
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label LblComCapLancCodForn 
      Caption         =   "Cod. Forn"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "FrmComCapLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
