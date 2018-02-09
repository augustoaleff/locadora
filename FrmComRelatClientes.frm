VERSION 5.00
Begin VB.Form FrmComRelatClientes 
   Caption         =   "Relatório de Clientes"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Visualizar em Tela"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton CmdComEmissaoConsultaPesquisaNome 
      Caption         =   "Pesquisa por Nome"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Período"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmComRelatClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()

End Sub
