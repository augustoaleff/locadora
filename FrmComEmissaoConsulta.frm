VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Pedidos"
   ClientHeight    =   7575
   ClientLeft      =   7740
   ClientTop       =   2175
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   13020
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   330
      Left            =   1320
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton CmdPesquisaNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   10440
      TabIndex        =   18
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   5
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   11040
      TabIndex        =   17
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame FrmDetalharPedido 
      Caption         =   "Detalhar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   6120
      Width           =   5415
      Begin VB.CommandButton CmdDetalhar 
         Caption         =   "Detalhar"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtNumeroPedido 
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
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label LblNumeroPedido 
         Caption         =   "Nª Pedido"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame FrmStatusAluguel 
      Caption         =   "Status Aluguel"
      Height          =   735
      Left            =   5520
      TabIndex        =   12
      Top             =   960
      Width           =   3975
      Begin VB.OptionButton OptAguardandoDevolucao 
         Caption         =   "Aguardando Devolução"
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton OptDevolvidos 
         Caption         =   "Devolvidos"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPedidos 
      Height          =   4095
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.TextBox TxtCodCliente 
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
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   330
      Left            =   3120
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
   Begin VB.Label LblValorTotal 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10560
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label LblTotal 
      Caption         =   "Total R$"
      Height          =   255
      Left            =   9720
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód . Cliente"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   945
   End
End
Attribute VB_Name = "FrmComEmissaoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultar_Click()
Dim VTOTAL As Double


 Dim DATADEV As Date
 Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
    
    If IsNumeric(TxtCodCliente.Text) <> Empty Then
    
    If Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = "" Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = "" Then
    
        If OptDevolvidos.Value = True Then
    
        reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " and Status = 'DEVOLVIDO' order by dataentrega")
    
        Else
    
            If OptAguardandoDevolucao.Value = True Then
    
            reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " and Status = 'ALUGADO' order by dataentrega")
    
            Else
        
            reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " order by dataentrega")
        
            End If
            
        End If
        

    Else
    
         If OptDevolvidos.Value = True Then
    
        reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " and dataentrega between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and Status = 'DEVOLVIDO' order by dataentrega")
    
        Else
    
            If OptAguardandoDevolucao.Value = True Then
    
            reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " and dataentrega between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and Status = 'ALUGADO' order by dataentrega")
    
            Else
        
            reg.Open ("SELECT * FROM PEDIDOS WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " and dataentrega between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' order by dataentrega")
        
            End If
            
        End If
    
    End If
    
    Call formata_flex
    
    Do Until reg.EOF = True
    
    If reg.Fields("DataDev") = "01/01/1900" Then
    
    MSFlexPedidos.AddItem (reg.Fields("NumPed") & vbTab & _
                          reg.Fields("CodVend") & vbTab & _
                          reg.Fields("Codcli") & vbTab & _
                          Format(reg.Fields("DataEntrega"), "DD/MM/YYYY") & vbTab & _
                          Format(reg.Fields("DataLimdev"), "DD/MM/YYYY") & vbTab & _
                          "" & vbTab & _
                          Format(reg.Fields("valorT"), "#,##0.00") & vbTab & _
                          Format(reg.Fields("valorp"), "#,##0.00") & vbTab & _
                          reg.Fields("status"))
                          
    Else
    
    MSFlexPedidos.AddItem (reg.Fields("NumPed") & vbTab & _
                          reg.Fields("CodVend") & vbTab & _
                          reg.Fields("Codcli") & vbTab & _
                          Format(reg.Fields("DataEntrega"), "DD/MM/YYYY") & vbTab & _
                          Format(reg.Fields("DataLimdev"), "DD/MM/YYYY") & vbTab & _
                          Format(reg.Fields("DataDev"), "DD/MM/YYYY") & vbTab & _
                          Format(reg.Fields("valorT"), "#,##0.00") & vbTab & _
                          Format(reg.Fields("valorp"), "#,##0.00") & vbTab & _
                          reg.Fields("status"))
                          
    End If
    
    reg.MoveNext
 
    Loop
    
    reg.Close
    
    VTOTAL = 0
    
    
    For contador = 1 To MSFlexPedidos.Rows - 1
    
    VTOTAL = VTOTAL + MSFlexPedidos.TextMatrix(contador, 6)
    
    Next
    
    LblValorTotal.Caption = Format(VTOTAL, "#,##0.00")
    
    Else
    
    MsgBox "Digite o Código do Cliente", vbInformation, Aviso
    
    End If
    
End Sub

Private Sub CmdDetalhar_Click()
 Dim NUMPED As Long

 If IsNumeric(TxtNumeroPedido.Text) <> Empty Then
 
 NUMPED = TxtNumeroPedido.Text
 FrmComEmissaoConsultaPedido.TxtNumeroPedido.Text = NUMPED
 FrmComEmissaoConsultaPedido.TxtNumeroPedido_KeyPress (13)
 
 End If
 


End Sub

Private Sub CmdLimparTela_Click()

TxtCodCliente.Text = ""
LblCliente.Caption = ""
MskPeriodoDe.Mask = ""
MskPeriodoDe.Text = ""
MskPeriodoDe.Mask = "##/##/####"
MskPeriodoAte.Mask = ""
MskPeriodoAte.Text = ""
MskPeriodoAte.Mask = "##/##/####"
OptTodos.Value = True
LblValorTotal.Caption = ""

Call formata_flex

TxtNumeroPedido.Text = ""

TxtCodCliente.SetFocus

End Sub

Private Sub CmdPesquisaNome_Click()

 FrmComEmissaoConsultaBuscarNome.Show

End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

 Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
    
    reg.Open ("SELECT NOME FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodCliente.Text) & "")
    
    If reg.EOF = False Then
    
    LblCliente.Caption = reg.Fields("Nome")
    
    MskPeriodoDe.SetFocus
    
    Else
    
    MsgBox "Cliente não encontrado"
    
    End If
    
    reg.Close
    
 End If
     
End Sub


Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) = True Or Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty) Then

 MskPeriodoAte.SetFocus
 
 End If
 
End Sub
Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) = True Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty) Then

 CmdConsultar.SetFocus
 
 End If
 
End Sub
Private Sub formata_flex()

MSFlexPedidos.Clear
MSFlexPedidos.Cols = 9
MSFlexPedidos.Rows = 1

MSFlexPedidos.Col = 0
MSFlexPedidos.Text = "Nº Ped."
MSFlexPedidos.ColWidth(0) = 1000

MSFlexPedidos.Col = 1
MSFlexPedidos.Text = "Cod.Vend."
MSFlexPedidos.ColWidth(1) = 1000

MSFlexPedidos.Col = 2
MSFlexPedidos.Text = "Cod.Cliente"
MSFlexPedidos.ColWidth(2) = 1000

MSFlexPedidos.Col = 3
MSFlexPedidos.Text = "Data Entrega"
MSFlexPedidos.ColWidth(3) = 1500


MSFlexPedidos.Col = 4
MSFlexPedidos.Text = "Data Lim. Dev."
MSFlexPedidos.ColWidth(4) = 1500

MSFlexPedidos.Col = 5
MSFlexPedidos.Text = "Data Devolução"
MSFlexPedidos.ColWidth(5) = 1500


MSFlexPedidos.Col = 6
MSFlexPedidos.Text = "Valor Total R$"
MSFlexPedidos.ColWidth(6) = 1500

MSFlexPedidos.Col = 7
MSFlexPedidos.Text = "Valor Pago R$"
MSFlexPedidos.ColWidth(7) = 1500

MSFlexPedidos.Col = 8
MSFlexPedidos.Text = "Status"
MSFlexPedidos.ColWidth(8) = 1500

End Sub



