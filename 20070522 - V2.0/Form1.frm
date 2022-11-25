VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGeraLinha 
   Caption         =   "Gerador de linhas de comando do EDGE - v 2.0 de 23/05/2007"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir o arquivo de Layout (.csv)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gerar Comandos do EDGE"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Analisar String de envio"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sair do programa"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
      Begin VB.CheckBox chkComent 
         Caption         =   "Incluir linha de Comentário"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Text            =   "@Res"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Variável de origem dos Text-Extract:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Posição Inicial do @POS:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3720
      Width           =   10575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Código"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Form1.frx":0000
         Left            =   360
         List            =   "Form1.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   5535
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":010B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2601
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Comandos do EDGE Gerados:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   1200
      TabIndex        =   10
      Top             =   720
      Width           =   9495
   End
   Begin VB.Label Label3 
      Caption         =   "Arquivo:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1095
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

If Label1(1).Caption = "" Then
   Label1(1).Caption = AbreDialogo()
End If
   
If Label1(1).Caption <> "" Then
   Open Label1(1).Caption For Input As #1
   
' Picture-format
If Combo1.ListIndex = -1 Then
   MsgBox "Selecione o tipo de código a ser gerado.", vbCritical, "Atenção"
End If

If Combo1.ListIndex = 0 Then
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
If Combo1.ListIndex = 1 Then
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
If Combo1.ListIndex = 2 Then
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
If Combo1.ListIndex = 3 Then
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
If Combo1.ListIndex = 4 Then
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
If Combo1.ListIndex = 5 Then
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
If Combo1.ListIndex = 6 Then
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
If Combo1.ListIndex = 7 Then
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       Label1(1).Caption = AbreDialogo()
    ElseIf Button.Index = 2 Then
       Command1_Click
    ElseIf Button.Index = 3 Then
       frmAnalisaString.Show
    ElseIf Button.Index = 4 Then
       End
    End If
End Sub
