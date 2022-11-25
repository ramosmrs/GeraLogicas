VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGeraLinha 
   Caption         =   "Gerador de linhas de comando do EDGE - v 1.5 de 23/05/2007"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Analisar String"
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opções"
      Height          =   2295
      Left            =   9600
      TabIndex        =   13
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkComent 
         Caption         =   "Incluir linha de Comentário"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Text            =   "@Res"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Variável de origem dos Text-Extract:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Posição Inicial do @POS:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar Comandos"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir arquivo"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1575
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2520
      Width           =   12375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Código"
      Height          =   2295
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton Option1 
         Caption         =   "Trim nos campos"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text-extract nos campos (valor fixo)"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text-extract em variáveis (valor fixo)"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture-format com campos"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "String de envio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text-extract em variáveis (usando @POS)"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text-extract nos campos (usando @POS)"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Picture-format"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmGeraLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim cont, x, i As Integer
Dim vLinha, tipo, just As String
Dim regs() As String
Dim linhanova As String

On Error GoTo trataerro
nrlinha = 1
Text1.Text = ""

If Text2.Text = "" Then
   Text2.Text = AbreDialogo()
End If
   
If Text2.Text <> "" Then
   Open Text2.Text For Input As #1
   
' Picture-format
If Option1(0).Value = True Then
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      If campo(1) = "A" Or campo(1) = "X" Then
         tipo = "#"
         just = "LEFT"
      ElseIf campo(1) = "N" Or campo(1) = "9" Then
         tipo = "%"
         just = "RIGHT"
      End If
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If

      linhanova = "ALWAYS PICTURE-FORMAT format value of """" using format picture " + Chr(34) + tipo + campo(2) + Chr(34) + " justification " + just + " store result in @" + campo(0)
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      nrlinha = nrlinha + 1
   Loop
   
End If

' Picture-format com campos
If Option1(4).Value = True Then
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      n = n + 1
      If campo(1) = "A" Or campo(1) = "X" Then
         tipo = "#"
         just = "LEFT"
      ElseIf campo(1) = "N" Or campo(1) = "9" Then
         tipo = "%"
         just = "RIGHT"
      End If
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If
      
      If campo(3) = "" Then
         linhanova = "ALWAYS PICTURE-FORMAT format value of """" using format picture " + Chr(34) + tipo + campo(2) + Chr(34) + " justification " + just + " store result in @" + campo(0)
      Else
         linhanova = "ALWAYS PICTURE-FORMAT format value of " + campo(3) + " using format picture " + Chr(34) + tipo + campo(2) + Chr(34) + " justification " + just + " store result in @" + campo(0)
      End If
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      
      nrlinha = nrlinha + 1
   Loop
End If

'Text-extract em variáveis (usando @POS)
If Option1(2).Value = True Then
   If Text3.Text = "" Then
      pos = "1"
   Else
      pos = LTrim(RTrim(Text3.Text))
   End If
   
   If Text4.Text = "" Then
      variavel = "@res"
   Else
      variavel = LTrim(RTrim(Text4.Text))
   End If
   
   Text1.Text = "ALWAYS COPY the value " + pos + " into @POS" + Chr(13) + Chr(10)
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If
      
      linhanova = "ALWAYS TEXT-EXTRACT using the value " + variavel + " starting at character position @POS extract " + campo(2) + " characters store result in @" + campo(0)
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      linhanova = "ALWAYS ADD the value of " + campo(2) + " to the value of @POS store result in @POS"
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      
      nrlinha = nrlinha + 1
   Loop
End If

'Text-extract em variáveis (usando valor fixo)
If Option1(5).Value = True Then
   If Text3.Text = "" Then
      n = 1
   Else
      n = Text3.Text * 1
End If

If Text4.Text = "" Then
   variavel = "@res"
Else
   variavel = LTrim(RTrim(Text4.Text))
End If
   
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))

      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If

      linhanova = "ALWAYS TEXT-EXTRACT using the value " + variavel + " starting at character position" + Str(n) + " extract " + campo(2) + " characters store result in @" + campo(0)
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      n = n + campo(2)
      
      nrlinha = nrlinha + 1
   Loop
End If

' Text-extract nos campos (usando @POS)
If Option1(1).Value = True Then
   If Text4.Text = "" Then
      variavel = "@res"
   Else
      variavel = LTrim(RTrim(Text4.Text))
   End If
   
   Text1.Text = "ALWAYS COPY the value 1 into @POS" + Chr(13) + Chr(10)
   
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      
      If campo(3) = "" Then
         linhanova = "ALWAYS TEXT-EXTRACT using the value " + variavel + " starting at character position @POS extract " + campo(2) + " characters store result in @" + campo(0)
      Else
         linhanova = "ALWAYS TEXT-EXTRACT using the value " + variavel + " starting at character position @POS extract " + campo(2) + " characters store result in " + campo(3)
      End If
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If
      
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      linhanova = "ALWAYS ADD the value of " + campo(2) + " to the value of @POS store result in @POS"
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      
      nrlinha = nrlinha + 1
   Loop
End If

'Text-extract em campos (usando valor fixo)
If Option1(6).Value = True Then
   If Text4.Text = "" Then
      variavel = "@res"
   Else
      variavel = LTrim(RTrim(Text4.Text))
   End If
   
   If Text3.Text = "" Then
      n = 1
   Else
      n = Text3.Text * 1
End If
   
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If
      
      linhanova = "ALWAYS TEXT-EXTRACT using the value " + variavel + " starting at character position" + Str(n) + " extract " + campo(2) + " characters store result in " + campo(3)
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      n = n + campo(2)
      
      nrlinha = nrlinha + 1
   Loop
End If

' String de envio
If Option1(3).Value = True Then
   cont = 0
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      campo(3) = LTrim(RTrim(campo(3)))
      If cont = 0 Then
         linhanova = "ALWAYS CALCULATE {CONCAT(@" + campo(0)
      ElseIf (cont Mod 5) = 0 Then
         linhanova = linhanova + ")} store result in @MSG.ACTX" + Chr(13) + Chr(10) + "ALWAYS CALCULATE {CONCAT(@MSG.ACTX, @" + campo(0)
      Else
         linhanova = linhanova + ", @" + campo(0)
      End If
      cont = cont + 1
      
      nrlinha = nrlinha + 1
   Loop
   
   linhanova = linhanova + ")} store result in @MSG.ACTX"
   Text1.Text = linhanova
   
End If

'Trim nos campos
If Option1(7).Value = True Then
   Do While Not EOF(1)
      Line Input #1, vLinha
      campo = Split(vLinha, ";")
      campo(0) = LTrim(RTrim(campo(0)))
      campo(0) = Replace(campo(0), " ", "_", , , vbTextCompare)
      campo(1) = LTrim(RTrim(campo(1)))
      campo(2) = LTrim(RTrim(campo(2))) * 1
      tamanho = UBound(campo)
      If tamanho = 4 Then
        campo(3) = LTrim(RTrim(campo(3)))
      Else
        campo(3) = ""
      End If
      
      If campo(3) = "" Then
        CampoTabela = "@" + campo(0)
      Else
        CampoTabela = campo(3)
      End If
      
      If campo(4) <> "" And chkComent.Value = 1 Then
         Text1.Text = Text1.Text + "--" + LTrim(RTrim(campo(4))) + Chr(13) + Chr(10)
      End If
      
      linhanova = "ALWAYS TRIM command EXTRA spaces from the value of " + CampoTabela + " store result in " + CampoTabela
      Text1.Text = Text1.Text + linhanova + Chr(13) + Chr(10)
      n = n + campo(2)
      
      nrlinha = nrlinha + 1
   Loop
End If

'-----------------------------------------------------------------
End If

trataerro:
   If Err.Number = 9 Then
       msgerro = "Ocorreu um erro na linha" + Str(nrlinha) + "."
       erro = MsgBox(msgerro, vbOKOnly, "Atenção")
       Text1.Text = ""
   End If

Close #1
End Sub


Private Sub Command2_Click()
   End
End Sub

Private Sub Command3_Click()
   Text2.Text = AbreDialogo()
End Sub

Function AbreDialogo() As String
   Dim prop As String
   CommonDialog1.DialogTitle = "Procurar Arquivos .csv"
   CommonDialog1.InitDir = App.Path
   prop = "Arquivos .csv |*.csv"
   CommonDialog1.Filter = prop
   CommonDialog1.FilterIndex = 1
   CommonDialog1.Flags = cdlOFNFileMustExist + _
                         cdlOFNLongNames + _
                         cdlOFNExplorer
   CommonDialog1.CancelError = False
   CommonDialog1.ShowOpen
   AbreDialogo = CommonDialog1.FileName
End Function

Sub trataerro()
   erro = MsgBox("Ocorreu um erro na linha" + nrlinha + ".", vbCritical, "Atenção", , 0)
End Sub

Private Sub Command4_Click()
    frmAnalisaString.Show
End Sub
