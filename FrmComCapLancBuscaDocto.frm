VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCapLancBuscarDocto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca por Nº Documento"
   ClientHeight    =   6795
   ClientLeft      =   17355
   ClientTop       =   5145
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdComCapLancBuscarDoctoLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CheckBox CBoxComCapLancDoctoAleatorio 
      Caption         =   "Aleatório"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdComCapLancBuscarDoctoBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox TxtComCapLancBuscarDoctoDescricao 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexComCapLancBuscarDoctoForn 
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
Attribute VB_Name = "FrmComCapLancBuscarDocto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
