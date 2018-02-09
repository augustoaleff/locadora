VERSION 5.00
Begin VB.Form FrmComCarLanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento Contas a Receber"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7875
   Begin VB.CommandButton CmdComCarLancBuscarDocto 
      Caption         =   "Busca por Nº Documento"
      Height          =   375
      Left            =   2640
      TabIndex        =   33
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdComCarLancBuscarNome 
      Caption         =   "Busca por Nome"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox TxtComCarLancSeq 
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   4800
      TabIndex        =   26
      Top             =   1680
      Width           =   2895
      Begin VB.TextBox TxtComCarLancValor 
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
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtComCarLancDesconto 
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
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtComCarLancJuros 
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
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtComCarLancValorTotal 
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LblComCarLancValor 
         Caption         =   "Valor R$"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblComCarLancDesconto 
         Caption         =   "Desconto R$"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LblComCarLancJuros 
         Caption         =   "Juros/Multa R$"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LblComCarLancValorTotal 
         Caption         =   "Valor Total R$"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdComCarLancLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton CmdComCarLancGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox TxtComCarLancNumeroDocto 
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
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox TxtComCarLancOBS 
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
      Left            =   600
      TabIndex        =   14
      Top             =   3720
      Width           =   7095
   End
   Begin VB.TextBox TxtComCarLancDataPagamento 
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
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox TxtComCarLancStatus 
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
      Left            =   720
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.ComboBox CmbComCarLancTipoDocto 
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox TxtComCarLancDataLancamento 
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
      Left            =   6120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtComCarLancNumeroPedido 
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
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox TxtComCarLancVencimento 
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
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox TxtComCarLancCliente 
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
      TabIndex        =   16
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox TxtComCarLancCodCliente 
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
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label LblComCarLancSeq 
      Caption         =   "Seq."
      Height          =   375
      Left            =   3480
      TabIndex        =   31
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label LblComCarLancNumeroDocto 
      Caption         =   "Nº Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LblComCarLancOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label LblComCarLancStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label LblComCarLancDataPagamento 
      Caption         =   "Data Pagamento"
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label LblComCarLancNumeroPedido 
      Caption         =   "Nº Pedido"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label LblComCarLancDataLancamento 
      Caption         =   "Data Lançamento"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LblComCarLancTipoDocto 
      Caption         =   "Tipo Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label LblComCarLancVencimento 
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblComCarLancCodCliente 
      Caption         =   "Cód. Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "FrmComCarLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label7_Click()

End Sub

