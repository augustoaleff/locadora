VERSION 5.00
Begin VB.Form FrmComCadProdutosNovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos"
   ClientHeight    =   5385
   ClientLeft      =   1515
   ClientTop       =   5670
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7965
   Begin VB.CommandButton CmdPesquisaDescricao 
      Caption         =   "Pesquisa por Descrição"
      Height          =   615
      Left            =   3360
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdNovoCodigo 
      Caption         =   "Criar Novo Código"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtCodBarras 
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
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   4455
   End
   Begin VB.TextBox TxtPreco 
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
      Left            =   3960
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox TxtCodForn 
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdCadastrarFoto 
      Caption         =   "Cadastrar Foto"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   5760
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
   End
   Begin VB.PictureBox PctFoto 
      Height          =   1815
      Left            =   6000
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   16
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame FrameEstoque 
      Caption         =   "Quantidade em Estoque"
      Height          =   1335
      Left            =   5880
      TabIndex        =   14
      Top             =   2880
      Width           =   2055
      Begin VB.TextBox TxtDisponivel 
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
         Left            =   1200
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxtAlugados 
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
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LblQuantDisponivel 
         Caption         =   "Disponível"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblAlugados 
         Caption         =   "Alugados"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
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
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
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
      ItemData        =   "FrmComCadProdutos.frx":0000
      Left            =   1200
      List            =   "FrmComCadProdutos.frx":0013
      TabIndex        =   4
      Top             =   1920
      Width           =   3495
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
      ItemData        =   "FrmComCadProdutos.frx":004B
      Left            =   1200
      List            =   "FrmComCadProdutos.frx":0058
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
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
      Height          =   330
      Left            =   1200
      MaxLength       =   70
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox TxtCodigo 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LblForn 
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
      TabIndex        =   23
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label LblCodBaras 
      Caption         =   "Cód. Barras"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label LblPreco 
      Caption         =   "Preço Aluguel R$"
      Height          =   495
      Left            =   3000
      TabIndex        =   21
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label LblCodForn 
      Caption         =   "Cod.Forn"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label LblLocalizacao 
      Caption         =   "Localização"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label LblCategoria 
      Caption         =   "Categoria"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label LblTipo 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label LblDescricao 
      Caption         =   "Descricão"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblCodigo 
      Caption         =   "Código"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmComCadProdutosNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdGravar_Click()
If ValidaCampos = True Then

Set CN1 = New ADODB.Connection
CN1.Open STR_DSN
Set reg = New ADODB.Recordset
reg.ActiveConnection = CN1
    
    reg.Open ("SELECT * FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodigo.Text) & "")
          
    'USO O INSERT
    
    If reg.EOF = True Then
    
     TxtDisponivel.Text = 0
     TxtAlugados.Text = 0
    

     CN1.Execute ("INSERT INTO PRODUTOS(CodProd,Descricao,Tipo,Categoria,Localizacao,Preco,CodForn,CodBarras,QuantEst,QuantAlug,Usuario,DataCad) " & _
     "VALUES (" & Trim(TxtCodigo.Text) & ",'" & StrConv(Trim(TxtDescricao.Text), vbUpperCase) & "','" & StrConv(Trim(CmbTipo.Text), vbUpperCase) & "','" & _
     StrConv(Trim(CmbCategoria.Text), vbUpperCase) & "','" & StrConv(Trim(TxtLocalizacao.Text), vbUpperCase) & "'," & Replace(Replace(Replace(Replace(Format(Trim(TxtPreco.Text), "#,##0.00"), "R", ""), "$", ""), " ", ""), ",", ".") & "," & _
     Trim(TxtCodForn.Text) & ",'" & Trim(TxtCodBarras.Text) & "','" & Trim(TxtDisponivel.Text) & "','" & Trim(TxtAlugados.Text) & "','','" & Format(Now, "YYYYMMDD hh:mm") & "')")
      
      MsgBox "Produto Cadastro com Sucesso", vbInformation, "Aviso"
      
    'USO O UPDATE
    Else

     CN1.Execute ("UPDATE PRODUTOS SET Descricao='" & StrConv(Trim(TxtDescricao.Text), vbUpperCase) & "',Tipo='" & StrConv(Trim(CmbTipo.Text), vbUpperCase) & "',Categoria='" & _
     StrConv(Trim(CmbCategoria.Text), vbUpperCase) & "',Localizacao='" & StrConv(Trim(TxtLocalizacao.Text), vbUpperCase) & "',Preco='" & Replace(Replace(Replace(Replace(Format(Trim(TxtPreco.Text), "#,##0.00"), "R", ""), "$", ""), " ", ""), ",", ".") & "',CodForn=" & _
     Trim(TxtCodForn.Text) & ",CodBarras='" & Trim(TxtCodBarras.Text) & "',QuantEst='" & Trim(TxtDisponivel.Text) & "',QuantAlug='" & Trim(TxtAlugados.Text) & _
     "' WHERE CODPROD = " & Trim(TxtCodigo.Text) & " ")
     
      MsgBox "Produto Atualizado com Sucesso", vbInformation, "Aviso"
      
    End If
       
    Call limpa_campos
    
    reg.Close
    
 End If
End Sub

Private Sub CmdLimparTela_Click()

Call limpa_campos


End Sub

Private Sub CmdNovoCodigo_Click()

 If TxtCodigo.Enabled = True Then
 
 Dim QUERY As String
 
 Set CN1 = New ADODB.Connection
 CN1.Open STR_DSN
 Set reg = New ADODB.Recordset
 reg.ActiveConnection = CN1
 
 
 CN1.Execute ("begin transaction")
 QUERY = "select UltCodProd from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltCodProd = UltCodProd+1;COMMIT"
 reg.Open (QUERY)
 
 TxtCodigo.Text = reg.Fields("UltCodProd")
 TxtCodigo.Enabled = False
 TxtDescricao.SetFocus
 
 reg.Close
 
 Else
 
  MsgBox "Limpe a Tela Antes de Criar um Novo Código", vbExclamation, "Aviso"
 
 End If
End Sub

Private Sub CmdPesquisaDescricao_Click()
 FrmComCadProdutosNovoPesquisa.Show
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 1000
End Sub
Private Sub limpa_campos()

 TxtCodigo.Enabled = True
 
 TxtCodigo.Text = ""
 TxtDescricao.Text = ""
 CmbTipo.Text = ""
 CmbCategoria.Text = ""
 TxtLocalizacao.Text = ""
 TxtPreco.Text = ""
 TxtCodForn.Text = ""
 LblForn.Caption = ""
 TxtCodBarras.Text = ""
 TxtDisponivel.Enabled = True
 TxtDisponivel.Text = ""
 TxtAlugados.Enabled = True
 TxtAlugados.Text = ""
 
 TxtAlugados.Enabled = False
 TxtDisponivel.Enabled = False
 
 TxtCodigo.SetFocus
End Sub
Private Function ValidaCampos() As Boolean

 If IsNumeric(TxtCodigo.Text) <> Empty Then

    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
 
 End If
 
 If TxtDescricao.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If CmbTipo.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If CmbCategoria.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If TxtLocalizacao.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If Format(TxtPreco.Text, "#,##0.00") <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If IsNumeric(TxtCodForn.Text) <> Empty And LblForn.Caption <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If TxtCodBarras.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If

End Function


Public Sub TxtCodForn_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsNumeric(TxtCodForn.Text) Then

Set CN1 = New ADODB.Connection
CN1.Open STR_DSN
Set REG2 = New ADODB.Recordset
REG2.ActiveConnection = CN1
 
REG2.Open ("SELECT RazaoSocial FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodForn.Text) & "")
 
 If REG2.EOF = False Then
 
 LblForn.Caption = REG2.Fields("RazaoSocial")
 
 TxtCodBarras.SetFocus
 
 Else
 
 MsgBox "Fornecedor não Encontrado", vbExclamation, "Aviso"
 TxtCodForn.SetFocus
 
 End If
 
 REG2.Close
 
 End If

End Sub

Public Sub TxtCodigo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsNumeric(TxtCodigo.Text) <> Empty Then

 Set CN1 = New ADODB.Connection
     CN1.Open STR_DSN
     Set reg = New ADODB.Recordset
     reg.ActiveConnection = CN1
 
     reg.Open ("SELECT * FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodigo.Text) & "")
 
 If reg.EOF = False Then
 
 TxtCodigo.Enabled = False
 TxtDisponivel.Enabled = True
 TxtAlugados.Enabled = True
 
 TxtDescricao.Text = reg.Fields("Descricao")
 CmbTipo.Text = reg.Fields("Tipo")
 CmbCategoria.Text = reg.Fields("Categoria")
 TxtLocalizacao.Text = reg.Fields("Localizacao")
 TxtPreco.Text = Format(reg.Fields("Preco"), "#,##0.00")
 TxtCodForn.Text = reg.Fields("CodForn")
 TxtCodForn_KeyPress (13)
 TxtCodBarras.Text = reg.Fields("CodBarras")
 TxtAlugados.Text = reg.Fields("QuantAlug")
 TxtDisponivel.Text = reg.Fields("QuantEst")
 
 reg.Close
 
 TxtDisponivel.Enabled = False
 TxtAlugados.Enabled = False
 TxtDescricao.SetFocus
 
 Else
 
 MsgBox "Código Não Existe", vbExclamation, "Aviso"
 CmdNovoCodigo.SetFocus
 
 End If
    
 End If

End Sub
Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtDescricao.Text <> Empty Then
 
 CmbTipo.SetFocus
 
 End If

End Sub
Private Sub CmbTipo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And CmbTipo.Text <> Empty Then
 
 CmbCategoria.SetFocus
 
 End If

End Sub
Private Sub CmbCategoria_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And CmbCategoria.Text <> Empty Then
 
 TxtLocalizacao.SetFocus
 
 End If

End Sub
Private Sub TxtLocalizacao_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtLocalizacao.Text <> Empty Then
 
 TxtPreco.SetFocus
 
 End If

End Sub
Private Sub TxtPreco_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Format(TxtPreco.Text, "#,##0.00") <> Empty Then
 
 TxtCodForn.SetFocus
 
 End If

End Sub
Private Sub TxtCodBarras_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtCodBarras.Text <> Empty Then
 
 CmdGravar.SetFocus
 
 End If

End Sub

