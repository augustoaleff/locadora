VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCadProdutosEnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Produtos"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MskDataEnt 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdBuscarForn 
      Caption         =   "Busca por Fornecedor"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton CmdBuscarDescricao 
      Caption         =   "Busca por Descrição"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox TxtNumeroNfe 
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
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox TxtQuantEnt 
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
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
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProduto 
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
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblDataEnt 
      Caption         =   "Data Entrada"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label LblEstoque 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label LblFornecedor 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   12
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label LblProduto 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   11
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label LblEstoqueAtual 
      Caption         =   "Estoque Atual"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label LblNumeroNfe 
      Caption         =   "Nº NFe"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label LblQuantEnt 
      Caption         =   "Quant. Entrada"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label LblCodForn 
      Caption         =   "Cód. Forn"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LblCodProduto 
      Caption         =   "Cód Produto"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "FrmComCadProdutosEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscarDescricao_Click()
 FrmComCadProdutosEntPesquisaDescricao.Show
End Sub

Private Sub CmdBuscarForn_Click()
 FrmComCadProdutosEntPesquisaForn.Show
End Sub

Private Sub CmdGravar_Click()
Dim ENTRADA As Long

If ValidaCampos = True Then

Set CN1 = New ADODB.Connection
CN1.Open STR_DSN

     CN1.Execute ("INSERT INTO ENTRADAS(CodProd,CodForn,QuantEnt,DataEnt,NumNfe) " & _
     "VALUES(" & Trim(TxtCodProduto.Text) & "," & Trim(TxtCodForn.Text) & ",'" & CInt(TxtQuantEnt.Text) & "','" & _
     Format(MskDataEnt.Text, "YYYYMMDD") & "','" & Trim(TxtNumeroNfe.Text) & "')")
    
    
     CN1.Execute ("update PRODUTOS set QuantEst = QuantEst + " & CInt(TxtQuantEnt.Text) & " where codProd = " & Trim(TxtCodProduto.Text) & "")
      
      
    MsgBox "Entrada Realizada com Sucesso", vbInformation, "Aviso"

       
    Call limpa_campos
    
 End If
End Sub

Private Sub CmdLimparTela_Click()
 Call limpa_campos
End Sub

Private Sub Form_Load()
 Me.Top = 1000
 Me.Left = 1000
End Sub

Private Sub limpa_campos()

 TxtCodProduto.Text = ""
 TxtCodForn.Text = ""
 LblProduto.Caption = ""
 LblFornecedor.Caption = ""
 TxtQuantEnt.Text = ""
 TxtNumeroNfe.Text = ""
 LblEstoque.Caption = ""
 MskDataEnt.Mask = ""
 MskDataEnt.Text = ""
 MskDataEnt.Mask = "##/##/####"
 
 TxtCodProduto.SetFocus

End Sub
Private Function ValidaCampos() As Boolean

 If IsNumeric(TxtCodProduto.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If IsNumeric(TxtCodForn.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
  If IsNumeric(TxtQuantEnt.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
  If IsNumeric(TxtCodProduto.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If TxtNumeroNfe.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If

 If IsDate(MskDataEnt.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
End Function

Public Sub TxtCodProduto_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsNumeric(TxtCodProduto.Text) <> Empty Then
 
Set CN1 = New ADODB.Connection
CN1.Open STR_DSN
Set REG2 = New ADODB.Recordset
REG2.ActiveConnection = CN1
 
REG2.Open ("SELECT DESCRICAO FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodProduto.Text) & "")
 
 If REG2.EOF = False Then
 
 LblProduto.Caption = REG2.Fields("Descricao")
 
 TxtCodForn.SetFocus
 
Set CN1 = New ADODB.Connection
CN1.Open STR_DSN
Set REG3 = New ADODB.Recordset
REG3.ActiveConnection = CN1

 
 REG3.Open ("SELECT QUANTEST FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodProduto.Text) & "")
 
 If REG3.EOF = False Then
 
 LblEstoque.Caption = REG3.Fields("QuantEst")
 
 End If
 
 REG3.Close
 
 Else
 
 MsgBox "Produto não Encontrado", vbExclamation, "Aviso"
 TxtCodProduto.SetFocus
 
 End If
 
 REG2.Close

 End If
 
End Sub

Public Sub TxtCodForn_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsNumeric(TxtCodProduto.Text) <> Empty Then
 
Set CN1 = New ADODB.Connection
CN1.Open STR_DSN
Set REG2 = New ADODB.Recordset
REG2.ActiveConnection = CN1
 
REG2.Open ("SELECT RAZAOSOCIAL FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodForn.Text) & "")
 
 If REG2.EOF = False Then
 
 LblFornecedor.Caption = REG2.Fields("RazaoSocial")
 
 TxtQuantEnt.SetFocus
 
 Else
 
 MsgBox "Produto não Encontrado", vbExclamation, "Aviso"
 TxtCodForn.SetFocus
 
 End If
 
 REG2.Close
 
 End If
End Sub

