VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCadForn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedores"
   ClientHeight    =   8205
   ClientLeft      =   600
   ClientTop       =   1290
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10335
   Begin VB.TextBox TxtCNPJ 
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
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
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
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
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
      Left            =   1320
      TabIndex        =   14
      Top             =   4680
      Width           =   2415
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
      Left            =   1320
      TabIndex        =   12
      Top             =   4200
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
      Left            =   6840
      TabIndex        =   13
      Top             =   4200
      Width           =   3135
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
      Left            =   9000
      TabIndex        =   11
      Top             =   3600
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
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   15
      Top             =   4680
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
      TabIndex        =   16
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdPesquisaNome 
      Caption         =   "Pesquisa por Nome"
      Height          =   615
      Left            =   3960
      TabIndex        =   30
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton CmdNovoCodigo 
      Caption         =   "Criar Novo Código"
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox TxtContato 
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
      Left            =   8520
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
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
      Left            =   1320
      TabIndex        =   0
      Top             =   330
      Width           =   1095
   End
   Begin VB.TextBox TxtRazaoSocial 
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
      Top             =   960
      Width           =   6615
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   7920
      TabIndex        =   19
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton CmdPesquisarCEP 
      Caption         =   "Pesquisar por CEP"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
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
      Left            =   1320
      TabIndex        =   17
      Top             =   5160
      Width           =   4335
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
      Height          =   1215
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   5640
      Width           =   8535
   End
   Begin VB.TextBox TxtIE 
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
      Left            =   4920
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox TxtRepresentante 
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
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin MSMask.MaskEdBox MskCelular 
      Height          =   345
      Left            =   4560
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
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   1320
      TabIndex        =   8
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
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label LblCEP 
      Caption         =   "CEP"
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label LblCelular 
      Caption         =   "Celular"
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label LblEndereco 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label LblComplemento 
      Caption         =   "Complemento"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label LblUF 
      Caption         =   "UF"
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   375
      Left            =   6120
      TabIndex        =   33
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label LblNumero 
      Caption         =   "Nº"
      Height          =   375
      Left            =   8640
      TabIndex        =   32
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label LblCodMunicipio 
      Caption         =   "Cod.Municipio"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label LblContato 
      Caption         =   "Contato"
      Height          =   255
      Left            =   7800
      TabIndex        =   28
      Top             =   2400
      Width           =   855
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
      Left            =   360
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LblRazaoSocial 
      Caption         =   "Razão Social"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblEmail 
      Caption         =   "E-mail"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label LblCNPJ 
      Caption         =   "CNPJ"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label LblIE 
      Caption         =   "IE"
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label LblRepresentante 
      Caption         =   "Representante"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Width           =   495
   End
End
Attribute VB_Name = "FrmComCadForn"
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



        reg.Open ("SELECT * FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodigo.Text) & "")

        'USO O INSERT
        If reg.EOF = True Then

            CN1.Execute ("INSERT INTO FORNECEDORES(CodForn,RazaoSocial,Representante,CNPJ,IE,Telefone,Celular,Contato,CEP,Endereco,NumEnd,Complemento,Bairro,Cidade,CodMun,UF,Email,OBS,Usuario,DataCad) " & _
                         "VALUES(" & Trim(TxtCodigo.Text) & ", '" & StrConv(Trim(TxtRazaoSocial.Text), vbUpperCase) & "','" & StrConv(Trim(TxtRepresentante.Text), vbUpperCase) & "', '" & _
                         Replace(Replace(Replace(Replace(Replace(TxtCNPJ.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "','" & StrConv(Trim(TxtIE.Text), vbUpperCase) & "','" & _
                         Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), "(", ""), ")", "") & "','" & _
                         Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), "(", ""), ")", "") & "','" & _
                         StrConv(Trim(TxtContato.Text), vbUpperCase) & "','" & Replace(Replace(Replace(Replace(Replace(MskCEP.Text, "-", ""), " ", ""), ".", ""), "/", ""), "\", "") & "', '" & _
                         StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "','" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "','" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "', '" & _
                         StrConv(Trim(TxtBairro.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & "','" & StrConv(Trim(TxtCodMunicipio.Text), vbUpperCase) & "','" & _
                         StrConv(Trim(TxtUF.Text), vbUpperCase) & "','" & StrConv(Trim(TxtEmail.Text), vbLowerCase) & "','" & StrConv(Trim(TxtOBS.Text), vbUpperCase) & "','','" & _
                         Format(Now, "YYYYMMDD hh:mm") & "') ")

            MsgBox "Fornecedor Cadastro com Sucesso", vbInformation, "Aviso"
            'USO O UPDATE
        Else

            CN1.Execute ("UPDATE FORNECEDORES SET RazaoSocial = '" & StrConv(Trim(TxtRazaoSocial.Text), vbUpperCase) & "', Representante='" & StrConv(Trim(TxtRepresentante.Text), vbUpperCase) & "', CNPJ = '" & _
                         Replace(Replace(Replace(Replace(Replace(TxtCNPJ.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "") & "',IE = '" & StrConv(Trim(TxtIE.Text), vbUpperCase) & "',Telefone='" & _
                         Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), "(", ""), ")", "") & "',Celular='" & _
                         Replace(Replace(Replace(Replace(Replace(Replace(Replace(MskCelular.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", ""), "(", ""), ")", "") & "',Contato='" & _
                         StrConv(Trim(TxtContato.Text), vbUpperCase) & "',CEP='" & Replace(Replace(Replace(Replace(Replace(MskCEP.Text, "-", ""), " ", ""), ".", ""), "/", ""), "\", "") & "',Endereco= '" & _
                         StrConv(Trim(TxtEndereco.Text), vbUpperCase) & "',NumEnd='" & StrConv(Trim(TxtNumero.Text), vbUpperCase) & "',Complemento='" & StrConv(Trim(TxtComplemento.Text), vbUpperCase) & "', Bairro='" & _
                         StrConv(Trim(TxtBairro.Text), vbUpperCase) & "',Cidade='" & StrConv(Trim(TxtCidade.Text), vbUpperCase) & "',CodMun='" & StrConv(Trim(TxtCodMunicipio.Text), vbUpperCase) & "',UF='" & _
                         StrConv(Trim(TxtUF.Text), vbUpperCase) & "',Email='" & StrConv(Trim(TxtEmail.Text), vbLowerCase) & "',OBS='" & StrConv(Trim(TxtOBS.Text), vbUpperCase) & "' " & _
                       " WHERE CODFORN = " & Trim(TxtCodigo.Text) & "")


            MsgBox "Fornecedor Atualizado com Sucesso", vbInformation, "Aviso"

        End If

        Call LimpaCampos

        reg.Close

    Else

        MsgBox "Por favor Verifique os Campos", vbInformation, "Aviso"

    End If


End Sub

Private Sub CmdLimparTela_Click()
    Call LimpaCampos
End Sub

Private Sub CmdNovoCodigo_Click()

    If TxtCodigo.Enabled = True Then

        Dim QUERY As String

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        CN1.Execute ("begin transaction")
        QUERY = "select UltCodForn from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltCodForn = UltCodForn+1;COMMIT"
        reg.Open (QUERY)

        TxtCodigo.Text = reg.Fields("UltCodForn")
        TxtCodigo.Enabled = False
        TxtRazaoSocial.SetFocus

        reg.Close

    Else

        MsgBox "Limpe a Tela Antes de Criar um Novo Código", vbExclamation, "Aviso"

    End If

End Sub

Private Sub CmdPesquisaNome_Click()
    FrmComCadFornPesquisa.Show
End Sub


Private Sub Form_Load()
    Me.Left = 1000
    Me.Top = 1000
End Sub
Private Function ValidaCampos() As Boolean

    If IsNumeric(TxtCodigo.Text) <> Empty Then

        ValidaCampos = True

        If TxtRazaoSocial.Text <> Empty Then

            ValidaCampos = True

            If TxtRepresentante.Text <> Empty Then

                ValidaCampos = True

                If Len(Replace(Replace(Replace(Replace(TxtCNPJ.Text, ".", ""), "-", ""), "/", ""), "\", "")) = 14 Then

                    ValidaCampos = True

                    If (Replace(Replace(Replace(Replace(TxtIE.Text, ".", ""), "-", ""), "/", ""), "\", "")) <> Empty Then

                        ValidaCampos = True

                        If Replace(Replace(Replace(Replace(Replace(MskTelefone.Text, "-", ""), "(", ""), ")", ""), " ", ""), ".", "") <> Empty Then

                            ValidaCampos = True

                            If Replace(Replace(Replace(Replace(Replace(MskCelular.Text, "-", ""), "(", ""), ")", ""), " ", ""), ".", "") <> Empty Then

                                ValidaCampos = True

                                If Replace(Replace(Replace(Replace(Replace(MskCEP.Text, "-", ""), "(", ""), ")", ""), " ", ""), ".", "") <> Empty Then

                                    ValidaCampos = True

                                Else

                                    ValidaCampos = False

                                End If

                            Else

                                ValidaCampos = False

                            End If

                        Else

                            ValidaCampos = False

                        End If

                    Else

                        ValidaCampos = False

                    End If

                Else

                    ValidaCampos = False

                End If

            Else

                ValidaCampos = False

            End If

        Else

            ValidaCampos = False

        End If

    Else

        ValidaCampos = False

    End If

End Function

Private Sub LimpaCampos()

    TxtCodigo.Enabled = True

    TxtCodigo.Text = ""
    TxtRazaoSocial.Text = ""
    TxtRepresentante.Text = ""
    TxtCNPJ.Text = ""
    TxtIE.Text = ""
    MskTelefone.Mask = ""
    MskTelefone.Text = ""
    MskTelefone.Mask = "(##) ####-####"
    MskCelular.Mask = ""
    MskCelular.Text = ""
    MskCelular.Mask = "(##) #####-####"
    TxtContato.Text = ""
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


Public Sub TxtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodigo.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT * FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodigo.Text) & "")

        If reg.EOF = False Then

            TxtCodigo.Enabled = False


            TxtRazaoSocial.Text = reg.Fields("RazaoSocial")
            TxtRepresentante.Text = reg.Fields("Representante")
            TxtCNPJ.Text = Format(reg.Fields("CNPJ"), "@@.@@@.@@@/@@@@-@@")
            TxtIE.Text = reg.Fields("IE")
            MskTelefone.Text = Format(reg.Fields("Telefone"), "(@@) @@@@-@@@@")
            MskCelular.Text = Format(reg.Fields("Celular"), "(@@) @@@@@-@@@@")
            TxtContato.Text = reg.Fields("Contato")
            MskCEP.Text = Format(reg.Fields("CEP"), "@@@@@-@@@")
            TxtEndereco.Text = reg.Fields("Endereco")
            TxtNumero.Text = reg.Fields("NumEnd")
            TxtComplemento.Text = reg.Fields("Complemento")
            TxtBairro.Text = reg.Fields("Bairro")
            TxtCidade.Text = reg.Fields("Cidade")
            TxtCodMunicipio.Text = reg.Fields("CodMun")
            TxtUF.Text = reg.Fields("UF")
            TxtEmail.Text = reg.Fields("Email")
            TxtOBS.Text = reg.Fields("OBS")

            'x = Format(VALOR, "#,##0.00")

            reg.Close

        Else

            MsgBox "Código Não Existe", vbExclamation, "Aviso"
            CmdNovoCodigo.SetFocus

        End If

    End If
End Sub

Private Sub TxtRazaoSocial_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtRazaoSocial.Text <> Empty Then

        TxtRepresentante.SetFocus
    End If

End Sub
Private Sub TxtRepresentante_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtRepresentante.Text <> Empty Then

        TxtCNPJ.SetFocus

    End If

End Sub
Private Sub TxtCNPJ_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(Replace(Replace(Replace(Replace(Replace(TxtCNPJ.Text, ".", ""), "-", ""), "/", ""), "\", ""), " ", "")) <> Empty Then

        TxtIE.SetFocus

    End If

End Sub
Private Sub TxtIE_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtIE.Text <> Empty Then

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

        TxtContato.SetFocus

    End If


End Sub
Private Sub TxtContato_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtContato.Text <> Empty Then

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


