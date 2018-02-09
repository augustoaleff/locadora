VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCadClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   8370
   ClientLeft      =   13050
   ClientTop       =   2175
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MskCelular 
      Height          =   345
      Left            =   4800
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(##) #####-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTelefone 
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(##) ####-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskCEP 
      Height          =   330
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####-###"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskDataNasc 
      Height          =   330
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.CommandButton CmdNovoCod 
      Caption         =   "Criar Novo Código"
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton CmdConsultaNome 
      Caption         =   "Consulta por Nome"
      Height          =   375
      Left            =   4080
      TabIndex        =   38
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox TxtUF 
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
      Left            =   9000
      MaxLength       =   2
      TabIndex        =   14
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox TxtCodMunicipio 
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
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox TxtNumero 
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
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton CmdCarregarFoto 
      Caption         =   "Carregar Foto"
      Height          =   375
      Left            =   8160
      TabIndex        =   34
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox PctFoto 
      Height          =   1935
      Left            =   7920
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   33
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox TxtRG 
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
      Left            =   5280
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TxtCPF 
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   19
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox TxtOBS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   5640
      Width           =   8655
   End
   Begin VB.TextBox TxtEmail 
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
      MaxLength       =   50
      TabIndex        =   15
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox TxtBairro 
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
      Left            =   6720
      MaxLength       =   50
      TabIndex        =   11
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox TxtComplemento 
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
      MaxLength       =   50
      TabIndex        =   10
      Top             =   4200
      Width           =   4575
   End
   Begin VB.CommandButton CmdPesquisarCEP 
      Caption         =   "Pesquisar por CEP"
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox TxtCidade 
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
      MaxLength       =   50
      TabIndex        =   12
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   360
      TabIndex        =   23
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   7920
      TabIndex        =   17
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox TxtEndereco 
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
      MaxLength       =   190
      TabIndex        =   8
      Top             =   3600
      Width           =   6975
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
      Height          =   330
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
      Width           =   6135
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
      MaxLength       =   18
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LblDataNasc 
      Caption         =   "Data Nasc./ Fundação"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LblCodMunicipio 
      Caption         =   "Cod.Municipio"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label LblNumero 
      Caption         =   "Nº"
      Height          =   375
      Left            =   8640
      TabIndex        =   36
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label LblCelular 
      Caption         =   "Celular"
      Height          =   255
      Left            =   4080
      TabIndex        =   35
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label LblRG 
      Caption         =   "IE"
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label LblCPF 
      Caption         =   "CPF/CNPJ"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblEmail 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label LblUF 
      Caption         =   "UF"
      Height          =   255
      Left            =   8640
      TabIndex        =   27
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label LblComplemento 
      Caption         =   "Complemento"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label LblCEP 
      Caption         =   "CEP"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label LblEndereco 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label LblNome 
      Caption         =   "Nome/Razão Social"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   1215
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
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmComCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultaNome_Click()
 FrmComCadClientesPesquisa.Show
End Sub

Private Sub CmdGravar_Click()

If ValidaCampos = True Then

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
 
    reg.Open ("SELECT CLIENTES.*,GETDATE() AS DATA FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodigo.Text) & "")
          
    'USO O INSERT
    If reg.EOF = True Then
 
      CN1.Execute ("INSERT INTO CLIENTES(CodCli,Nome,CPF,IE,DataNasc,Telefone,Celular,CEP,Endereco,NumEnd,Complemento,Bairro,Cidade,CodMun,UF,Email,Obs,Usuario,DataCad) " & _
      "VALUES(" & Trim(TxtCodigo.Text) & ",'" & StrConv(TxtNome.Text, vbUpperCase) & "','" & _
      Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", "") & "','" & StrConv(Trim(TxtRG.Text), vbUpperCase) & "','" & Format(MskDataNasc.Text, "YYYYMMDD") & "','" & Replace(Replace(Replace(Replace(Trim(MskTelefone.Text), "(", ""), ")", ""), "-", ""), " ", "") & "', '" & _
      Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", "") & "','" & Replace(Replace(MskCEP.Text, "-", ""), " ", "") & "', '" & StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "', '" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "','" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "','" & _
      StrConv(Trim(TxtBairro.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCodMunicipio.Text), vbUpperCase) & "','" & Trim(TxtUF.Text) & "','" & StrConv(TxtEmail.Text, vbLowerCase) & "','" & StrConv(TxtOBS.Text, vbUpperCase) & "','','" & Format(reg.Fields("DATA"), "yyyymmdd hh:mm") & "')")
      
      MsgBox "Cliente Cadastro com Sucesso", vbInformation, "Aviso"
    'USO O UPDATE
    Else

      CN1.Execute ("UPDATE CLIENTES SET Nome = '" & StrConv(TxtNome.Text, vbUpperCase) & "',CPF =  '" & Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", "") & "',IE = '" & StrConv(Trim(TxtRG.Text), vbUpperCase) & "',DataNasc= '" & _
      Format(MskDataNasc.Text, "YYYYMMDD") & "',Telefone = '" & Replace(Replace(Replace(Replace(MskTelefone.Text, "(", ""), ")", ""), "-", ""), " ", "") & "',Celular = '" & Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", "") & "',CEP = '" & _
      Replace(Replace(MskCEP.Text, "-", ""), " ", "") & "',Endereco = '" & StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "',NumEnd = '" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "',Complemento = '" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "',Bairro = '" & StrConv(Trim(TxtBairro.Text), vbUpperCase) & "',Cidade = '" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & _
      "' ,CodMun = " & Trim(TxtCodMunicipio.Text) & ",UF = '" & StrConv(Trim(TxtUF.Text), vbUpperCase) & "' ,Email = '" & StrConv(TxtEmail.Text, vbLowerCase) & "' ,Obs = '" & Trim(TxtOBS.Text) & "' " & _
      " WHERE CODCLI = " & Trim(TxtCodigo.Text) & "")
      
      MsgBox "Cliente Atualizado com Sucesso", vbInformation, "Aviso"
      
    End If
       
     
    
    Call limpa_campos
    
    reg.Close
 
 
 Else
 
 MsgBox "Por favor Verifique os Campos", vbInformation, "Aviso"
 
 
 End If
 
End Sub
Private Sub limpa_campos()

 TxtCodigo.Enabled = True

 TxtCodigo.Text = ""
 TxtNome.Text = ""
 TxtCPF.Text = ""
 TxtRG.Text = ""
 MskDataNasc.Mask = ""
 MskDataNasc.Text = ""
 MskDataNasc.Mask = "##/##/####"
 MskTelefone.Mask = ""
 MskTelefone.Text = ""
 MskTelefone.Mask = "(##) ####-####"
 MskCelular.Mask = ""
 MskCelular.Text = ""
 MskCelular.Mask = "(##) #####-####"
 MskCEP.Mask = ""
 MskCEP.Text = ""
 MskCEP.Mask = "#####-###"
 TxtEndereco.Text = ""
 TxtNumero.Text = ""
 TxtComplemento.Text = ""
 TxtBairro.Text = ""
 TxtCidade.Text = ""
 TxtCodMunicipio.Text = ""
 TxtUF.Text = ""
 TxtEmail.Text = ""
 TxtOBS.Text = ""
 
 TxtCodigo.SetFocus
 
End Sub
Private Function ValidaCampos() As Boolean

 If IsNumeric(TxtCodigo.Text) <> Empty Then
  
  ValidaCampos = True
  
 Else
  
  ValidaCampos = False
  
 End If
 
 
 If TxtNome.Text <> Empty Then
 
   ValidaCampos = True
   
 Else
    
   ValidaCampos = False
    
 End If
 
 
 If Len(Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", "")) = 11 Or Len(Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", "")) = 14 Then
 
   ValidaCampos = True
   
 Else
 
   ValidaCampos = False
   
 End If
 
 
 If IsDate(MskDataNasc.Text) = True Then
 
   ValidaCampos = True
   
 Else
   
   ValidaCampos = False
 
 End If
 
 
 If Replace(Replace(Replace(Replace(MskTelefone.Text, "(", ""), ")", ""), "-", ""), " ", "") <> Empty Then
    ValidaCampos = True
 Else
    ValidaCampos = False
 End If
 
  If Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", "") <> Empty Then
    ValidaCampos = True
 Else
    ValidaCampos = False
 End If
 
 
 If Replace(Replace(MskCEP.Text, "-", ""), " ", "") <> Empty Then
    ValidaCampos = True
 Else
    ValidaCampos = False
 End If
 
 
 If Trim(TxtEndereco.Text) <> Empty Or Trim(TxtBairro.Text) <> Empty Or Trim(TxtCidade.Text) <> Empty Or Trim(TxtCodMunicipio.Text) <> Empty Then
    ValidaCampos = True
 Else
    ValidaCampos = False
 End If
 
 
  
End Function

Private Sub CmdLimparTela_Click()
 Call limpa_campos
End Sub

Private Sub CmdNovoCod_Click()

 If TxtCodigo.Enabled = True Then
 
 Dim QUERY As String
 
 Set CN1 = New ADODB.Connection
 CN1.Open STR_DSN
 Set reg = New ADODB.Recordset
 reg.ActiveConnection = CN1
 
 
 CN1.Execute ("begin transaction")
 QUERY = "select UltCodCli from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltCodCli = UltCodCli+1;COMMIT"
 reg.Open (QUERY)
 
 TxtCodigo.Text = reg.Fields("UltCodCli")
 TxtCodigo.Enabled = False
 TxtNome.SetFocus
 
 reg.Close
 
 Else
 
  MsgBox "Limpe a Tela Antes de Criar um Novo Código", vbExclamation, "Aviso"
 
 End If
 
 
End Sub

Private Sub Form_Load()
 Me.Top = 1000
 Me.Left = 1000
End Sub


Public Sub TxtCodigo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsNumeric(TxtCodigo.Text) <> Empty Then

Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
 
    reg.Open ("SELECT * FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodigo.Text) & "")
 
 If reg.EOF = False Then
 
 TxtCodigo.Enabled = False
 
 TxtNome.Text = reg.Fields("Nome")
  If Len(reg.Fields("CPF")) = 11 Then
        TxtCPF.Text = Format(reg.Fields("CPF"), "@@@.@@@.@@@-@@")
  End If
  If Len(reg.Fields("CPF")) = 14 Then
        TxtCPF.Text = Format(reg.Fields("CPF"), "@@.@@@.@@@/@@@@-@@")
  End If
 TxtRG.Text = reg.Fields("IE")
 MskDataNasc.Text = Format(reg.Fields("DataNasc"), "DD/MM/YYYY")
 'MskTelefone.Mask = ""
 MskTelefone.Text = Format(reg.Fields("Telefone"), "(@@) @@@@-@@@@")
 'MskCelular.Mask = ""
 MskCelular.Text = Format(reg.Fields("Celular"), "(@@) @@@@@-@@@@")
 'MskCEP.Mask = ""
 MskCEP.Text = Format(reg.Fields("CEP"), "@@@@@-@@@")
 TxtEndereco.Text = reg.Fields("Endereco")
 TxtNumero.Text = reg.Fields("NumEnd")
 TxtComplemento.Text = reg.Fields("Complemento")
 TxtBairro.Text = reg.Fields("bairro")
 TxtCidade.Text = reg.Fields("cidade")
 TxtCodMunicipio.Text = reg.Fields("codmun")
 TxtUF.Text = reg.Fields("UF")
 TxtEmail.Text = reg.Fields("email")
 TxtOBS.Text = reg.Fields("obs")
 
 'x = Format(VALOR, "#,##0.00")
 
 reg.Close
 
 Else
 
 MsgBox "Código Não Existe", vbExclamation, "Aviso"
 CmdNovoCod.SetFocus
 
 End If
    
 End If

 
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtNome.Text <> Empty Then
  
  TxtCPF.SetFocus
 
 End If

End Sub
Private Sub TxtCPF_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtCPF.Text <> Empty Then
  
  TxtRG.SetFocus
 
 End If

End Sub
Private Sub TxtRG_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtRG.Text <> Empty Then
  
  MskDataNasc.SetFocus
 
 End If

End Sub
Private Sub MskDataNasc_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsDate(MskDataNasc.Text) = True Then
  MskTelefone.SetFocus
 
 End If

End Sub
Private Sub MskTelefone_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, "(", ""), ")", ""), "-", ""), " ", ""), "_", "")) = 10 Then
  '(11) 2514-2536
  MskCelular.SetFocus
 
 End If

End Sub
Private Sub MskCelular_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", ""), "_", "")) = 11 Then
  
  MskCEP.SetFocus
 
 End If
 

End Sub
Private Sub MskCEP_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(MskCEP.Text, "-", ""), "_", "")) = 8 Then
  
  CmdPesquisarCEP.SetFocus
 
 End If

End Sub
Private Sub TxtNumero_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtNumero.Text <> Empty Then
  
  TxtComplemento.SetFocus
 
 End If

End Sub
Private Sub TxtComplemento_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
  
  TxtEmail.SetFocus
  
 End If

End Sub
Private Sub TxtEmail_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtEmail.Text <> Empty Then
  
  TxtOBS.SetFocus
  
 End If

End Sub
Private Sub TxtOBS_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
  
  CmdGravar.SetFocus
  
 End If

End Sub


