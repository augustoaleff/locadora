VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComRelatProdutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Produtos"
   ClientHeight    =   3960
   ClientLeft      =   16170
   ClientTop       =   2535
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar Tela"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox TxtLocalizacao 
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox CmbCategoria 
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.ComboBox CmbTipo 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TxtDescricao 
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
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label LblDataCadastro 
      Caption         =   "Data Cadastro"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label LblDescricao 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LblLocalizacao 
      Caption         =   "Localização"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LblCategoria 
      Caption         =   "Categoria"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label LblTipo 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "FrmComRelatProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
