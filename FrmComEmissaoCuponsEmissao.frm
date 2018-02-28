VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoCuponsEmissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Cupom de Desconto"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6030
   Begin MSMask.MaskEdBox MskValidoDe 
      Height          =   330
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
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
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox TxtPorcentagem 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame FrameDesconto 
      Caption         =   "Desconto em"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton OptPorcentagem 
         Caption         =   "Porcentagem"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptValor 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdGerarCupom 
      Caption         =   "Gerar Cupom"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox TxtCodCupom 
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
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MskValidoAte 
      Height          =   330
      Left            =   3480
      TabIndex        =   17
      Top             =   2040
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
   Begin VB.TextBox TxtValor 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label LblPorcent 
      Caption         =   "%"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label LblAte 
      Caption         =   "até"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label LblValidoDe 
      Caption         =   "Válido de"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label LblRS 
      Caption         =   "R$"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label LblDescricao 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label LblDesconto 
      Caption         =   "Desconto:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label LblCodCupom 
      Caption         =   "Código Cupom"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "FrmComEmissaoCuponsEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIPO As String
Dim VALOR As Double
Private Sub limpa_campos()

    OptValor.Value = True
    TxtCodCupom.Text = ""
    TxtValor.Text = ""
    TxtPorcentagem.Text = ""
    MskValidoDe.Mask = ""
    MskValidoDe.Text = ""
    MskValidoDe.Mask = "##/##/####"
    MskValidoAte.Mask = ""
    MskValidoAte.Text = ""
    MskValidoAte.Mask = "##/##/####"
    TxtDescricao.Text = ""
    TxtCodCupom.SetFocus

End Sub

Private Sub CmdGerarCupom_Click()

    If ValidaCampos = True Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Dim DATADE, DATAATE As Date


        If OptValor.Value = True And TxtValor.Text <> Empty Then

            TIPO = "V"
            VALOR = CDbl(TxtValor.Text)
        Else
            If OptPorcentagem.Value = True And TxtPorcentagem.Text <> Empty Then

                TIPO = "P"
                VALOR = CDbl(TxtPorcentagem.Text) / 100
            End If
        End If
        
        If Replace(Replace(MskValidoDe.Text, "/", ""), "_", "") = Empty Then
        
            DATADE = "01/01/1900"
            
        Else
        
            DATADE = MskValidoDe.Text
            
        End If
        
         If Replace(Replace(MskValidoAte.Text, "/", ""), "_", "") = Empty Then
        
            DATAATE = "31/12/2199"
            
        Else
        
            DATAATE = MskValidoAte.Text
            
        End If
        
        
        

        CN1.Execute ("INSERT INTO CUPONS(CodCupom,Tipo,Valor,ValidadeDe,ValidadeAte,Descricao,Status,Usuario,DataEmissao)" & _
                     " VALUES('" & StrConv(Trim(TxtCodCupom.Text), vbUpperCase) & "','" & TIPO & "'," & Replace(VALOR, ",", ".") & ",'" & _
                     Format(DATADE, "YYYYMMDD") & "','" & Format(DATAATE, "YYYYMMDD") & "','" & Trim(StrConv(TxtDescricao.Text, vbUpperCase)) & "','NAOUTILIZADO','','" & Format(Now, "YYYYMMDD hh:mm") & "')")



        MsgBox "Cupom Criado", vbInformation, Criação
        Call limpa_campos

    Else

        MsgBox "Verifique os Campos", vbExclamation, Aviso

    End If


End Sub
Private Function ValidaCampos() As Boolean

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set REG2 = New ADODB.Recordset
    REG2.ActiveConnection = CN1

    If TxtCodCupom.Text <> Empty Then

        REG2.Open ("SELECT CODCUPOM FROM CUPONS WHERE CODCUPOM = '" & Trim(TxtCodCupom.Text) & "'")

        If REG2.EOF = True Then

            ValidaCampos = True

            If (OptValor.Value = True And TxtValor.Text <> Empty) Or (OptPorcentagem.Value = True And TxtPorcentagem.Text <> Empty) Then

                ValidaCampos = True

                If (IsDate(MskValidoDe.Text) Or Replace(Replace(MskValidoDe.Text, "/", ""), "_", "") = Empty) <> Empty Then

                    ValidaCampos = True

                    If (IsDate(MskValidoAte.Text) Or Replace(Replace(MskValidoAte.Text, "/", ""), "_", "") = Empty) <> Empty Then

                        ValidaCampos = True

                        If TxtDescricao.Text <> Empty Then

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

            MsgBox "Já existe um cupom com esse código", vbInformation, Aviso

        End If

    Else

        ValidaCampos = False

    End If

    REG2.Close


End Function


Private Sub CmdLimparTela_Click()

    Call limpa_campos

End Sub


Private Sub Form_Load()

    Me.Left = 1500
    Me.Top = 1500

End Sub

Private Sub OptPorcentagem_Click()

    If OptPorcentagem.Value = True Then
    
    TxtValor.Visible = False
    LblRS.Visible = False
    
    TxtPorcentagem.Visible = True
    LblPorcent.Visible = True
    
    End If
    
End Sub

Private Sub OptValor_Click()

    If OptValor.Value = True Then
    
    TxtValor.Visible = True
    LblRS.Visible = True
    
    TxtPorcentagem.Visible = False
    LblPorcent.Visible = False
    
    End If
    

End Sub

Private Sub TxtCodCupom_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtCodCupom.Text <> Empty Then

        If OptValor.Value = True Then

            TxtValor.SetFocus

        Else

            If OptPorcentagem.Value = True Then

                TxtPorcentagem.SetFocus
            End If
        End If
    End If

End Sub


Private Sub TxtValor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtValor.Text <> Empty Then

        MskValidoDe.SetFocus

    End If

End Sub
Private Sub TxtPorcentagem_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtPorcentagem.Text <> Empty Then

        MskValidoDe.SetFocus

    End If

End Sub
Private Sub MskValidoDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskValidoDe.Text) Or Replace(Replace(MskValidoDe.Text, "/", ""), "_", "") = Empty) Then

        MskValidoAte.SetFocus

    End If

End Sub
Private Sub MskValidoAte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskValidoAte.Text) Or Replace(Replace(MskValidoAte.Text, "/", ""), "_", "") = Empty) Then

        TxtDescricao.SetFocus

    End If

End Sub
Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtDescricao.Text <> Empty Then

        CmdGerarCupom.SetFocus

    End If

End Sub


