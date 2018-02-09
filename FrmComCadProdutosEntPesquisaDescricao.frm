VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCadProdutosEntPesquisaDescricao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca por Descricao"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CheckBox CBoxAleatorio 
      Caption         =   "Aleatório"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComCadProdutosEntPesquisaDescricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Me.Top = 1500
 Me.Left = 1500
End Sub
Private Sub CmdBuscar_Click()

If ValidaCampos() = True Then


    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
 
 

 If CBoxAleatorio.Value = vbChecked Then
      
      reg.Open ("SELECT codprod,descricao FROM PRODUTOS WHERE DESCRICAO LIKE '%" & Trim(TxtDescricao.Text) & "%'")
      
 Else
       
      reg.Open ("SELECT codprod,descricao FROM PRODUTOS WHERE DESCRICAO LIKE '" & Trim(TxtDescricao.Text) & "%'")
     
 End If

 
 Call formata_flex
 
 Do Until reg.EOF = True
    
  MSFlexPesquisa.AddItem (reg.Fields("codprod") & vbTab & _
                         reg.Fields("descricao"))
                         
  reg.MoveNext
    
 Loop
 
 
 reg.Close
 End If
 
End Sub
Private Sub formata_flex()
 
 MSFlexPesquisa.Clear
 MSFlexPesquisa.Cols = 2
 MSFlexPesquisa.Rows = 1

 MSFlexPesquisa.Col = 0
 MSFlexPesquisa.Text = "Cód."
 MSFlexPesquisa.ColWidth(0) = 700
 
 MSFlexPesquisa.Col = 1
 MSFlexPesquisa.Text = "Descricao"
 MSFlexPesquisa.ColWidth(1) = 5400
 
End Sub
Private Sub CmdLimparTela_Click()
 Call limpa_campos
End Sub
Private Sub limpa_campos()

    CBoxAleatorio.Value = vbUnchecked
    TxtDescricao.Text = ""
    Call formata_flex
    
    TxtDescricao.SetFocus
End Sub

Private Function ValidaCampos() As Boolean


If TxtDescricao.Text <> Empty Then

    ValidaCampos = True
    
Else

    ValidaCampos = False
    
End If


End Function

Private Sub MSFlexPesquisa_KeyPress(KeyAscii As Integer)
 
 Dim CODIGO As Long
 
 If KeyAscii = 13 Then
  
  MSFlexPesquisa.Col = 0
  CODIGO = Trim(MSFlexPesquisa.Text)
  
  FrmComCadProdutosEnt.TxtCodProduto.Text = CODIGO
  FrmComCadProdutosEnt.TxtCodProduto_KeyPress (13)
  Unload Me
  
 End If

End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtDescricao.Text <> Empty Then
    
    Call CmdBuscar_Click
    
    MSFlexPesquisa.SetFocus
    
    End If
    
End Sub
