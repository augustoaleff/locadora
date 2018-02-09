VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoAluguelBuscarNome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar por Nome"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6750
   Begin VB.Frame FrameBuscarPor 
      Caption         =   "Buscar Por"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton OptNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptCPF 
         Caption         =   "CPF/CNPJ"
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptRG 
         Caption         =   "RG/IE"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptCEP 
         Caption         =   "CEP"
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox TxtNome 
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox CBoxAleatorio 
      Caption         =   "Nome Aleatório"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComEmissaoAluguelBuscarNome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
