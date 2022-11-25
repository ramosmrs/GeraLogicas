VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnalisaString 
   Caption         =   "Analisador de  String"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   1200
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
            Picture         =   "frmAnalisaString.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalisaString.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalisaString.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalisaString.frx":24F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir arquivo de layout"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Analisar string"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Limpar o resultado da análise"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Voltar para a tela principal"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4320
      Width           =   13335
   End
   Begin VB.TextBox txtString 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1920
      Width           =   13335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTrabalhando 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Resultado da análise:  (ø = espaço)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3975
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
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   12255
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
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "String a ser analisada:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frmAnalisaString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim matriz(4) As String
    matriz(0) = "|"
    matriz(1) = "/"
    matriz(2) = "-"
    matriz(3) = "\"
    
   Dim varctrl As Integer

Private Sub Command1_Click()
Dim cont, x, i As Integer
Dim vLinha, tipo, just As String
Dim regs() As String
Dim linhanova As String

If txtString.Text = "" Then
   MsgBox "A String a ser verificada está vazia.", vbCritical, "Atenção"
   Exit Sub
End If

On Error GoTo trataerro
nrlinha = 1
txtResult.Text = ""

If Label1.Caption = "" Then
   Label1.Caption = AbreDialogo()
End If
   
If Label1.Caption <> "" Then
   Open Label1.Caption For Input As #1
   
   cont = 1
   Do While Not EOF(1)
      AtualizaTrabalhandoLabel
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
      
      'If campo(4) <> "" Then
      '   txtResult.Text = txtResult.Text + LTrim(RTrim(campo(4)))
      'End If

      linhanova = Format(nrlinha, "00000") + " - " + campo(0) + "(Pos.:" + Format(cont, "00000") + " - Tam.:" + campo(2) + ")" + " -> " + Replace(Mid(txtString.Text, cont, campo(2)), " ", Chr(248), , , vbTextCompare)
      
      cont = cont + campo(2)
      txtResult.Text = txtResult.Text + linhanova + Chr(13) + Chr(10)
      nrlinha = nrlinha + 1
   Loop
   
End If

trataerro:
   If Err.Number = 9 Then
       msgerro = "Ocorreu um erro na linha" + Str(nrlinha) + "."
       erro = MsgBox(msgerro, vbOKOnly, "Atenção")
       Text1.Text = ""
   End If

Close #1

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       Label1.Caption = AbreDialogo()
    ElseIf Button.Index = 2 Then
       Command1_Click
    ElseIf Button.Index = 3 Then
       txtResult = ""
    ElseIf Button.Index = 4 Then
       frmAnalisaString.Hide
    End If
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

Sub AtualizaTrabalhandoLabel()
   
   If varctrl = 3 Then
      varctrl = 0
   Else
      varctrl = varctrl + 1
   End If
   
   lblTrabalhando.Caption = matriz(varctrl)
   
End Sub


