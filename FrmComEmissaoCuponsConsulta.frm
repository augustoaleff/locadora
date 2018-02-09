VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComEmissaoCuponsConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cupons"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8595
   Begin VB.CommandButton Command3 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Relatório"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option3 
         Caption         =   "Não Utilizados"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Utilizados"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      Caption         =   "a"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   6360
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cód Cupom"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "FrmComEmissaoCuponsConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
