VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCadFunc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funcionários"
   ClientHeight    =   8970
   ClientLeft      =   12930
   ClientTop       =   3045
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10155
   Begin VB.TextBox TxtCargo 
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
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   19
      Top             =   6000
      Width           =   2415
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
      MaxLength       =   70
      TabIndex        =   16
      Top             =   5040
      Width           =   4335
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
      Top             =   840
      Width           =   6135
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
      TabIndex        =   9
      Top             =   3480
      Width           =   6975
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
      TabIndex        =   13
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton CmdPesquisarCEP 
      Caption         =   "Pesquisar por CEP"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
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
      TabIndex        =   11
      Top             =   4080
      Width           =   4575
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
      TabIndex        =   12
      Top             =   4080
      Width           =   3135
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
      Top             =   1320
      Width           =   2415
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
      Top             =   1320
      Width           =   2055
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
      TabIndex        =   10
      Top             =   3480
      Width           =   975
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
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
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
      TabIndex        =   15
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox TxtCodGerente 
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
      Left            =   6000
      TabIndex        =   20
      Top             =   6120
      Width           =   1095
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
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   6600
      Width           =   8775
   End
   Begin VB.CheckBox CBoxDesligado 
      Caption         =   "Desligado?"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   240
      Width           =   1095
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
      Width           =   1095
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   7440
      TabIndex        =   22
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   360
      TabIndex        =   27
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox PctFoto 
      Height          =   1935
      Left            =   8040
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   26
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdCarregarFoto 
      Caption         =   "Carregar Foto"
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton CmdConsultaNome 
      Caption         =   "Consulta por Nome"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton CmdNovoCod 
      Caption         =   "Criar Novo Código"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   240
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MskCelular 
      Height          =   345
      Left            =   4800
      TabIndex        =   6
      Top             =   2280
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
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
      Top             =   2880
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
      Top             =   1800
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
   Begin MSMask.MaskEdBox MskDataAdmissao 
      Height          =   330
      Left            =   1440
      TabIndex        =   17
      Top             =   5520
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
   Begin MSMask.MaskEdBox MskDataDemissao 
      Height          =   330
      Left            =   7440
      TabIndex        =   18
      Top             =   5520
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
   Begin VB.Label LblUF 
      Caption         =   "UF"
      Height          =   255
      Left            =   8640
      TabIndex        =   50
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label LblNumero 
      Caption         =   "Nº"
      Height          =   255
      Left            =   8520
      TabIndex        =   49
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label LblEmail 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   240
      TabIndex        =   48
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label LblNome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   840
      Width           =   495
   End
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label LblEndereco 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label LblCEP 
      Caption         =   "CEP"
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label LblComplemento 
      Caption         =   "Complemento"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   375
      Left            =   6000
      TabIndex        =   41
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label LblCPF 
      Caption         =   "CPF"
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LblRG 
      Caption         =   "RG"
      Height          =   255
      Left            =   4920
      TabIndex        =   39
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LblCelular 
      Caption         =   "Celular"
      Height          =   255
      Left            =   4080
      TabIndex        =   38
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label LblCodMunicipio 
      Caption         =   "Cod.Municipio"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label LblDataNasc 
      Caption         =   "Data Nasc."
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label LblGerente 
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
      Left            =   7200
      TabIndex        =   35
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label LblCodGerente 
      Caption         =   "Cód Gerente"
      Height          =   255
      Left            =   5040
      TabIndex        =   34
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label LbCargo 
      Caption         =   "Cargo"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label LblComCadFuncDataDemissao 
      Caption         =   "Data Demissão"
      Height          =   255
      Left            =   6240
      TabIndex        =   31
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label LblDataAdmissao 
      Caption         =   "Data Adimissão"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   5520
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
      TabIndex        =   28
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmComCadFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultaNome_Click()
 FrmComCadFuncPesquisa.Show
End Sub

Private Sub CmdGravar_Click()

If ValidaCampos = True Then

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1
    Dim DATA_DEMISSAO As String
    
    
    If TxtCodGerente.Text = "" Then
     TxtCodGerente.Text = "0"
    End If
        
    reg.Open ("SELECT * FROM FUNCIONARIOS WHERE CODFUNC = " & Trim(TxtCodigo.Text) & "")
          
    'USO O INSERT
    If reg.EOF = True Then
   
 
      CN1.Execute ("INSERT INTO FUNCIONARIOS(CodFunc,Nome,CPF,RG,DataNasc,Telefone,Celular,CEP,Endereco,NumEnd,Complemento,Bairro,Cidade,CodMun,UF,Email,DataAdm,DataDem,Cargo,CodGerente,OBS,Usuario,DataCad) " & _
      "VALUES (" & Trim(TxtCodigo.Text) & ",'" & StrConv(TxtNome.Text, vbUpperCase) & "','" & Replace(Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "','" & _
      Replace(Replace(Replace(Replace(Replace(TxtRG.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "','" & Format(MskDataNasc.Text, "YYYYMMDD") & "','" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "','" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "', '" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskCEP.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "','" & _
      StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "','" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "','" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "','" & _
      StrConv(Trim(TxtBairro.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCodMunicipio.Text), vbUpperCase) & "', '" & _
      StrConv(TxtUF.Text, vbUpperCase) & "','" & StrConv(TxtEmail.Text, vbLowerCase) & "','" & Format(MskDataAdmissao.Text, "YYYYMMDD") & "','" & Format(Replace(Replace(MskDataDemissao.Text, "_", ""), "/", ""), "YYYYMMDD") & "','" & _
      StrConv(TxtCargo.Text, vbUpperCase) & "'," & Trim(TxtCodGerente.Text) & ",'" & StrConv(TxtOBS.Text, vbUpperCase) & "','','" & Format(Now, "YYYYMMDD hh:mm") & "') ")
     
      
      MsgBox "Funcionário Cadastro com Sucesso", vbInformation, "Aviso"
    'USO O UPDATE
    Else

      CN1.Execute ("UPDATE FUNCIONARIOS SET Nome = '" & StrConv(TxtNome.Text, vbUpperCase) & "',CPF='" & Replace(Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "',RG = '" & _
      Replace(Replace(Replace(Replace(Replace(TxtRG.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "',DataNasc='" & Format(MskDataNasc.Text, "YYYYMMDD") & "',Telefone = '" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "',Celular = '" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "', CEP = '" & _
      Replace(Replace(Replace(Replace(Replace(Replace(MskCEP.Text, "(", ""), ")", ""), "-", ""), " ", ""), "/", ""), "\", "") & "',Endereco = '" & _
      StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "',NumEnd='" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "',Complemento='" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "',Bairro='" & _
      StrConv(Trim(TxtBairro.Text), vbUpperCase) & "',Cidade='" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & "',CodMun='" & StrConv(Trim(TxtCodMunicipio.Text), vbUpperCase) & "',UF= '" & _
      StrConv(TxtUF.Text, vbUpperCase) & "',Email='" & StrConv(TxtEmail.Text, vbLowerCase) & "',DataAdm='" & Format(MskDataAdmissao.Text, "YYYYMMDD") & "',DataDem='" & Format(Replace(Replace(MskDataDemissao.Text, "/", ""), "_", ""), "YYYYMMDD") & "',Cargo='" & _
      StrConv(TxtCargo.Text, vbUpperCase) & "',CodGerente = " & Trim(TxtCodGerente.Text) & ",OBS='" & StrConv(TxtOBS.Text, vbUpperCase) & "' WHERE CODFUNC = " & Trim(TxtCodigo.Text) & "")
    
      MsgBox "Funcionário Atualizado com Sucesso", vbInformation, "Aviso"
      
    End If
       
    Call limpa_campos
    
    reg.Close
 End If
End Sub

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
 QUERY = "select UltCodFunc from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltCodFunc = UltCodFunc+1;COMMIT"
 reg.Open (QUERY)
 
 TxtCodigo.Text = reg.Fields("UltCodFunc")
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
 MskDataAdmissao.Mask = ""
 MskDataAdmissao.Text = ""
 MskDataAdmissao.Mask = "##/##/####"
 MskDataDemissao.Mask = ""
 MskDataDemissao.Text = ""
 MskDataDemissao.Mask = "##/##/####"
 TxtCargo.Text = ""
 TxtCodGerente.Text = ""
 LblGerente.Caption = ""
 TxtOBS.Text = ""
 
 TxtCodigo.SetFocus
 
    
End Sub
Private Function ValidaCampos() As Boolean

 If IsNumeric(TxtCodigo) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 
 If TxtNome.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If Len(Replace(Replace(Replace(Replace(Replace(TxtCPF.Text, "-", ""), ".", ""), "/", ""), "\", ""), " ", "")) = 11 Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If Len(Replace(Replace(Replace(Replace(Replace(TxtRG.Text, "-", ""), ".", ""), "/", ""), "\", ""), " ", "")) = 9 Then
    
    ValidaCampos = True
 
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If IsDate(MskDataNasc.Text) <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False

 End If
 
 
 If Replace(Replace(Replace(Replace(MskTelefone.Text, "-", ""), "(", ""), ")", ""), " ", "") <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If Replace(Replace(Replace(Replace(MskCelular.Text, "-", ""), "(", ""), ")", ""), " ", "") <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If Replace(Replace(Replace(Replace(MskCEP.Text, "-", ""), "(", ""), ")", ""), " ", "") <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If TxtEndereco.Text <> Empty Or TxtNumero.Text <> Empty Or TxtBairro.Text <> Empty Then
 
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 
 If TxtCidade.Text <> Empty Or TxtCodMunicipio.Text <> Empty Or TxtUF.Text <> Empty Then
    
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If


If TxtEmail <> Empty Then

    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If IsDate(MskDataAdmissao.Text) <> Empty Then
    
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 
 If IsDate(MskDataDemissao.Text) = True Then
    
    ValidaCampos = True
    
 Else
 
    ValidaCampos = False
    
 End If
 
 If TxtCargo.Text <> Empty Then
 
    ValidaCampos = True
 
 Else
 
    ValidaCampos = False

 End If
 
 
End Function

Public Sub TxtCodigo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsNumeric(TxtCodigo.Text) <> Empty Then

 Set CN1 = New ADODB.Connection
     CN1.Open STR_DSN
     Set reg = New ADODB.Recordset
     reg.ActiveConnection = CN1
     Set REG2 = New ADODB.Recordset
     REG2.ActiveConnection = CN1
 
     reg.Open ("SELECT * FROM FUNCIONARIOS WHERE CODFUNC = " & Trim(TxtCodigo.Text) & "")
     
 
 If reg.EOF = False Then
 
 TxtCodigo.Enabled = False
 
 TxtNome.Text = reg.Fields("Nome")
 TxtCPF.Text = Format(reg.Fields("CPF"), "@@@.@@@.@@@-@@")
 TxtRG.Text = Format(reg.Fields("RG"), "@@.@@@.@@@-@")
 MskDataNasc.Text = Format(reg.Fields("DataNasc"), "DD/MM/YYYY")
 MskTelefone.Text = Format(reg.Fields("Telefone"), "(@@) @@@@-@@@@")
 MskCelular.Text = Format(reg.Fields("Celular"), "(@@) @@@@@-@@@@")
 MskCEP.Text = Format(reg.Fields("CEP"), "@@@@@-@@@")
 TxtEndereco.Text = reg.Fields("Endereco")
 TxtNumero.Text = reg.Fields("NumEnd")
 TxtComplemento.Text = reg.Fields("Complemento")
 TxtBairro.Text = reg.Fields("Bairro")
 TxtCidade.Text = reg.Fields("Cidade")
 TxtCodMunicipio.Text = reg.Fields("CodMun")
 TxtUF.Text = reg.Fields("UF")
 TxtEmail.Text = reg.Fields("Email")
 MskDataAdmissao.Text = Format(reg.Fields("DataAdm"), "DD/MM/YYYY")
 TxtCodGerente.Text = reg.Fields("CodGerente")
 TxtCodGerente_KeyPress (13)
 
 If Format(reg.Fields("DataDem"), "DD/MM/YYYY") = "01/01/1900" Then
     MskDataDemissao.Text = "__/__/____"
 Else
    MskDataDemissao.Text = Format(reg.Fields("DataDem"), "DD/MM/YYYY")
 End If
 
 TxtCargo.Text = reg.Fields("Cargo")
 TxtOBS.Text = reg.Fields("OBS")
 
 'x = Format(VALOR, "#,##0.00")
 
 reg.Close
 
 TxtNome.SetFocus
 
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

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(TxtCPF.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "")) = 11 Then

 TxtRG.SetFocus
 
 End If

End Sub
Private Sub TxtRG_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(TxtRG.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "")) = 8 Or Len(Replace(Replace(Replace(Replace(Replace(TxtRG.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "")) = 9 Then

 MskDataNasc.SetFocus
 
 End If

End Sub
Private Sub MskDataNasc_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsDate(MskDataNasc.Text) = True Then

 MskTelefone.SetFocus
 
 End If

End Sub
Private Sub MskTelefone_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), ")", ""), "(", ""), "_", "")) = 10 Then

 MskCelular.SetFocus
 
 End If

End Sub
Private Sub MskCelular_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), ")", ""), "(", ""), "_", "")) = 11 Then

 MskCEP.SetFocus
 
 End If

End Sub
Private Sub MskCEP_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskCEP.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), ")", ""), "(", ""), "_", "")) = 8 Then

 TxtEndereco.SetFocus
 
 End If

End Sub
Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtEndereco.Text <> Empty Then

 TxtNumero.SetFocus
 
 End If

End Sub
Private Sub TxtNumero_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtNumero.Text <> Empty Then

 TxtComplemento.SetFocus
 
 End If

End Sub
Private Sub TxtComplemento_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtComplemento.Text <> Empty Then

 TxtBairro.SetFocus
 
 End If

End Sub
Private Sub TxtBairro_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtBairro.Text <> Empty Then

 TxtCidade.SetFocus
 
 End If

End Sub
Private Sub TxtCidade_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtCidade.Text <> Empty Then

 TxtCodMunicipio.SetFocus
 
 End If

End Sub
Private Sub TxtCodMunicipio_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtCodMunicipio.Text <> Empty Then

 TxtUF.SetFocus
 
 End If

End Sub
Private Sub TxtUF_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtUF.Text <> Empty Then

 TxtEmail.SetFocus
 
 End If

End Sub
Private Sub TxtEmail_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtEmail.Text <> Empty Then

 MskDataAdmissao.SetFocus
 
 End If

End Sub
Private Sub MskDataAdmissao_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And IsDate(MskDataAdmissao.Text) = True Then

 TxtCargo.SetFocus
 
 End If

End Sub
Private Sub TxtCargo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtCargo.Text <> Empty Then

 TxtCodGerente.SetFocus
 
 End If

End Sub
Private Sub TxtCodGerente_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
 
 If TxtCodGerente.Text <> "" Then
 
 Set CN1 = New ADODB.Connection
     CN1.Open STR_DSN
     Set REG2 = New ADODB.Recordset
     REG2.ActiveConnection = CN1
    
   REG2.Open ("SELECT CodFunc,Nome FROM FUNCIONARIOS WHERE CODFUNC = " & Trim(TxtCodGerente.Text) & "")
   
  If REG2.EOF = False Then
   
  LblGerente.Caption = REG2.Fields("Nome")
  Else
  
 LblGerente.Caption = ""
 TxtCodGerente.Text = ""
  
  End If
   
   REG2.Close
 End If
 

 TxtOBS.SetFocus
 
 End If

End Sub
Private Sub TxtOBS_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtOBS.Text <> Empty Then

 CmdGravar.SetFocus
 
 End If

End Sub
