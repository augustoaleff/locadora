VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCarLancBuscarNome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar por Nome"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6765
   Begin VB.TextBox TxtComCarLancBuscarNomeDescricao 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton CmdComCarLancBuscarNomeBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox CBoxComCarLancBuscarNomeAleatorio 
      Caption         =   "Aleatório"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdComCarLancBuscarNomeLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComCarLancBuscarNomeForn 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComCarLancBuscarNome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
