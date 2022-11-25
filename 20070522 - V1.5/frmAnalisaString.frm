VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnalisaString 
   Caption         =   "Analisador de  String"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   1335
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
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4200
      Width           =   9975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir arquivo"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analisar String"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1335
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
      Top             =   1800
      Width           =   9975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   8
      Top             =   3840
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
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   8895
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
      TabIndex        =   6
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
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "frmAnalisaString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

      linhanova = Format(nrlinha, "0000") + " - " + campo(0) + "(" + campo(2) + ")" + " -> " + Replace(Mid(txtString.Text, cont, campo(2)), " ", Chr(248), , , vbTextCompare)
      
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

Private Sub Command2_Click()
    frmAnalisaString.Hide
End Sub

Private Sub Command3_Click()
    Label1.Caption = AbreDialogo()
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

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
   txtResult = ""
End Sub
