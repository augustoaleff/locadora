VERSION 5.00
Begin VB.Form FrmComEmissaoCuponsEmissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Cupom de Desconto"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6030
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoValidadeAte 
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
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoValidadeDe 
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoPorcentagem 
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
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame FrameComEmissaoCuponsEmissaoDesconto 
      Caption         =   "Desconto em"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton OptComEmissaoCuponsEmissaoPorcentagem 
         Caption         =   "Porcentagem"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptComEmissaoCuponsEmissaoValor 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdComEmissaoCuponsEmissaoGerarCupom 
      Caption         =   "Gerar Cupom"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoDescricao 
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
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoValor 
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
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox TxtComEmissaoCuponsEmissaoCodCupom 
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
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoAte 
      Caption         =   "até"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoValidoDe 
      Caption         =   "Válido de"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoOu 
      Caption         =   "ou"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoRS 
      Caption         =   "R$"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoPorcent 
      Caption         =   "%"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoDescricao 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoDesconto 
      Caption         =   "Desconto:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label LblComEmissaoCuponsEmissaoCodCupom 
      Caption         =   "Código Cupom"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "FrmComEmissaoCuponsEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

