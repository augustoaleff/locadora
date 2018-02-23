VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCarLanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento Contas a Receber"
   ClientHeight    =   5205
   ClientLeft      =   960
   ClientTop       =   2175
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7875
   Begin MSMask.MaskEdBox MskDataPagto 
      Height          =   330
      Left            =   3600
      TabIndex        =   24
      Top             =   3240
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
   Begin VB.TextBox TxtValorTotal 
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
      Left            =   5880
      TabIndex        =   23
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CmdBuscarDocto 
      Caption         =   "Busca por Nº Documento"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Busca por Nome"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox TxtSeq 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox TxtNumeroDocto 
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
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
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
      Height          =   330
      Left            =   600
      TabIndex        =   7
      Top             =   3720
      Width           =   7095
   End
   Begin VB.TextBox TxtStatus 
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
      Top             =   3240
      Width           =   495
   End
   Begin VB.ComboBox CmbTipoDocto 
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
      ItemData        =   "FrmComCarLanc.frx":0000
      Left            =   1440
      List            =   "FrmComCarLanc.frx":0002
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
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
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MskDataLancto 
      Height          =   330
      Left            =   6000
      TabIndex        =   25
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MskVencto 
      Height          =   330
      Left            =   1440
      TabIndex        =   26
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
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LblCliente 
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
      Left            =   2640
      TabIndex        =   21
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label LblSeq 
      Caption         =   "Seq."
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label LblNumeroDocto 
      Caption         =   "Nº Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Lbltatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label LblDataPagamento 
      Caption         =   "Data Pagamento"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label LblNumeroPedido 
      Caption         =   "Nº Pedido"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label LblDataLancamento 
      Caption         =   "Data Lançamento"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LblTipoDocto 
      Caption         =   "Tipo Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label LblVencimento 
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód. Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "FrmComCarLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscarDocto_Click()

    FrmComCarLancBuscarDocto.Show
    
End Sub

Private Sub CmdBuscarNome_Click()

    FrmComCarLancBuscarNome.Show
    
End Sub
Public Sub limpa_campos()

    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    MskVencto.Mask = ""
    MskVencto.Text = ""
    MskVencto.Mask = "##/##/####"
    TxtSeq.Text = ""
    MskDataLancto.Mask = ""
    MskDataLancto.Text = ""
    MskDataLancto.Mask = "##/##/####"
    TxtNumeroPedido.Text = ""
    TxtNumeroDocto.Text = ""
    TxtValorTotal.Text = ""
    CmbTipoDocto.Text = ""
    TxtStatus.Text = ""
    MskDataPagto.Mask = ""
    MskDataPagto.Text = ""
    MskDataPagto.Mask = "##/##/####"
    TxtOBS.Text = ""
    
    TxtCodCliente.SetFocus
    
    

End Sub
Private Function ValidaCampos() As Boolean

    If IsNumeric(TxtCodCliente.Text) <> Empty Then

        ValidaCampos = True

        If IsDate(MskVencto.Text) <> Empty Then

            ValidaCampos = True

            If IsNumeric(TxtSeq.Text) <> Empty Then

                ValidaCampos = True

                If IsDate(MskDataLancto.Text) <> Empty Then

                    ValidaCampos = True

                    If IsNumeric(TxtNumeroPedido.Text) <> Empty Or TxtNumeroPedido.Text = Empty Then

                        ValidaCampos = True

                        If TxtNumeroDocto.Text <> Empty Then

                            ValidaCampos = True

                            If IsNumeric(TxtValorTotal.Text) <> Empty Then

                                ValidaCampos = True

                                If CmbTipoDocto.Text <> Empty Then

                                    ValidaCampos = True

                                    If TxtStatus.Text <> Empty Then

                                        ValidaCampos = True

                                            If (TxtStatus.Text = "P" And Replace(Replace(MskDataPagto.Text, "/", ""), "_", "") <> Empty) Or (TxtStatus.Text = "A" And Replace(Replace(MskDataPagto.Text, "/", ""), "_", "") = Empty) Or (TxtStatus.Text = "C" And Replace(Replace(MskDataPagto.Text, "/", ""), "_", "") = Empty) Then

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

        Else

            ValidaCampos = False

        End If

    Else

        ValidaCampos = False

    End If



End Function


Private Sub CmdGravar_Click()

    Call ValidaCampos
    
    If ValidaCampos = True Then

        Dim PAGTO As Date
        
        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT CodCli,Vencto,Seq FROM C_A_R WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " AND Vencto = '" & Format(MskVencto.Text, "YYYYMMDD") & "' AND SEQ = " & Trim(TxtSeq.Text) & "")
        
        
        If StrConv(TxtStatus.Text, vbUpperCase) = "A" Or StrConv(TxtStatus.Text, vbUpperCase) = "C" Then
        
        PAGTO = "01/01/1900"
        
        Else
        
        PAGTO = Format(MskDataPagto.Text, "DD/MM/YYYY")
        
        End If
        

        If reg.EOF = True Then

            CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao)" & _
                         "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(MskVencto.Text, "YYYYMMDD") & "'," & Trim(TxtSeq.Text) & ",'" & Format(Now, "YYYYMMDD") & "'," & _
                         TxtNumeroPedido.Text & ",'" & StrConv(TxtNumeroDocto.Text, vbUpperCase) & "','" & StrConv(CmbTipoDocto.Text, vbUpperCase) & "','" & StrConv(Trim(TxtStatus.Text), vbUpperCase) & "','" & _
                         Format(PAGTO, "YYYYMMDD") & "','" & StrConv(OBS.Text, vbUpperCase) & "'," & Replace(Trim(CDbl(TxtValorTotal.Text)), ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

        Else

            CN1.Execute ("UPDATE C_A_R  SET DataLancto='" & Format(MskDataLancto, "YYYYMMDD") & "',NumPed=" & TxtNumeroPedido.Text & "," & _
                         "NumDocto = '" & StrConv(TxtNumeroDocto.Text, vbUpperCase) & "',Tipo='" & StrConv(CmbTipoDocto.Text, vbUpperCase) & "'," & _
                         "Status='" & StrConv(Trim(TxtStatus.Text), vbUpperCase) & "',DataPagto='" & Format(PAGTO, "YYYYMMDD") & "',OBS = '" & StrConv(OBS, vbUpperCase) & "'," & _
                         "Valor = " & Replace(Trim(CDbl(TxtValorTotal.Text)), ",", ".") & ",DataEmissao='" & Format(Now, "YYYYMMDD hh:mm") & "'" & _
                         "WHERE CodCli= " & Trim(TxtCodCliente.Text) & " AND Vencto = '" & Format(MskVencto.Text, "YYYYMMDD") & "' AND Seq=" & Trim(TxtSeq.Text) & "")
        End If
        
        Call limpa_campos

    Else

        MsgBox "Verifique os Campos", vbInformation, Aviso

    End If


End Sub

Private Sub CmdLimparTela_Click()
    
    Call limpa_campos
    
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
            MskVencto.SetFocus

        Else

            MsgBox "Cliente não Encontrado"

        End If

        reg.Close

    End If

End Sub
Private Sub MskVencto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskVencto.Text) Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1
        Dim SEQ As Integer

        reg.Open ("SELECT Seq FROM C_A_R WHERE VENCTO = '" & Format(MskVencto.Text, "YYYYMMDD") & "' and CODCLI = " & Trim(TxtCodCliente.Text) & " order by seq desc")

        If reg.EOF = True Then

            SEQ = 1


        Else

            SEQ = reg.Fields("seq") + 1


        End If

        reg.Close

        TxtSeq.Text = SEQ
        MskDataLancto = Format(Now, "DD/MM/YYYY")
        TxtNumeroPedido.SetFocus
        
    End If
    
    End Sub

Private Sub TxtNumeroPedido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsNumeric(TxtNumeroPedido.Text) <> Empty Or TxtNumeroPedido.Text = Empty) Then
    
        TxtNumeroDocto.SetFocus
        
    End If
    

End Sub
Private Sub TxtNumeroDocto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNumeroDocto.Text <> Empty Then
    
        TxtValorTotal.SetFocus
        
    End If
    

End Sub

Private Sub TxtValorTotal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtValorTotal.Text) <> Empty Then
    
        CmbTipoDocto.SetFocus
        TxtValorTotal = Format(TxtValorTotal, "#,##0.00")
        
    End If
    

End Sub
Private Sub CmbTipoDocto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbTipoDocto.Text <> Empty Then
    
        TxtStatus.SetFocus
        TxtStatus.Text = "A"
        
    End If
    

End Sub
Private Sub TxtStatus_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (StrConv(TxtStatus.Text, vbUpperCase) = "A" Or StrConv(TxtStatus.Text, vbUpperCase) = "C") Then
        
        TxtOBS.SetFocus
        MskDataPagto.Mask = ""
        MskDataPagto.Text = ""
        MskDataPagto.Mask = "##/##/####"
    
    ElseIf KeyAscii = 13 And StrConv(TxtStatus.Text, vbUpperCase) = "P" Then
        
        TxtOBS.SetFocus
        MskDataPagto.Text = Format(Now, "DD/MM/YYYY")
        
    ElseIf KeyAscii = 13 And TxtStatus <> Empty Then
    
        MsgBox "Digite somente Status P - Pago, A- Aberto ou C - Cancelado"
        
    End If
    
End Sub
Private Sub TxtOBS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        CmdGravar.SetFocus
        
    End If
    

End Sub


    




