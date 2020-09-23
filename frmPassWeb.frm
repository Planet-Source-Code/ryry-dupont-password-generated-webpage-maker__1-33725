VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPassWeb 
   Caption         =   "Webpage Password Encoder"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Encode It"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtC 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtB 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Password"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(case sensitive)"
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Save as:"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "File to encrypt:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click(Index As Integer)
    CD.FileName = ""
    If Index = 0 Then
        CD.ShowOpen
        If CD.FileName = "" Then Exit Sub
        txtB.Text = CD.FileName
    Else
        CD.ShowSave
        If CD.FileName = "" Then Exit Sub
        txtC.Text = CD.FileName
    End If
End Sub
Private Sub cmdDoIt_Click()
    Dim tmpAA, thePAGE, tmpBB, newPAGE
    If txtA.Text = "" Or txtB.Text = "" Or txtC.Text = "" Then
        MsgBox "All fields muct be filled in", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    Reset
    tmpAA = ""
    thePAGE = ""
    Open txtB.Text For Input As #1
        Do Until EOF(1)
            Line Input #1, tmpAA
            thePAGE = thePAGE & tmpAA
        Loop
    Close #1
    tmpAA = 1
    tmpBB = txtA.Text
    newPAGE = ""
    For z = 1 To Len(thePAGE)
        If Asc(Mid(thePAGE, z, 1)) < 32 Then Mid(thePAGE, z, 1) = " "
        newPAGE = newPAGE & Format(Asc(Mid(thePAGE, z, 1)) + Asc(Mid(tmpBB, tmpAA, 1)) - 32, "000")
        tmpAA = tmpAA + 1
        If tmpAA > Len(tmpBB) Then tmpAA = 1
    Next z
    Reset
    Open txtC.Text For Output As #1
        Print #1, "<HTML><BODY><SCRIPT SRC=" & Mid(txtC.Text, InStrRev(txtC.Text, "\") + 1, InStrRev(txtC.Text, ".") - InStrRev(txtC.Text, "\") - 1) & ".js></SCRIPT></BODY></HTML>"
    Close #1
    Reset
    Open Left(txtC.Text, InStrRev(txtC.Text, ".")) & "js" For Output As #1
        Print #1, "var content=" & Chr(34) & newPAGE & Chr(34)
        Print #1, "var pass=prompt(" & Chr(34) & "Enter Password" & Chr(34) & ")"
        Print #1, "var outtie = " & Chr(34) & Chr(34)
        Print #1, "var currpos = 0"
        Print #1, "var i = 0"
        Print #1, "var Key = " & Chr(34) & "  '#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_ `abcdefghijklmnopqrstuvwxyz{|}~ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ" & Chr(34)
        Print #1, "for (i=0; i<content.length; i=i+3) {"
        Print #1, "  outtie = outtie + Key.charAt((content.substring(i, i + 3) - Key.indexOf(pass.charAt(currpos))) - 32)"
        Print #1, "  currpos = currpos + 1"
        Print #1, "  if (currpos==pass.length) {"
        Print #1, "    currpos = 0"
        Print #1, "  }"
        Print #1, "}"
        Print #1, "document.writeln (outtie)"
    Close #1
    MsgBox "Wrote files:" & vbCrLf & vbTab & txtC.Text & vbCrLf & vbTab & Left(txtC.Text, InStrRev(txtC.Text, ".")) & "js"
End Sub
Private Sub Form_Load()
    CD.Filter = "Web Page Files(*.htm, *.html)|*.htm;*.html|All Files(*.*)|*.*"
End Sub
Private Sub txtA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtA.Text = "Password" Then txtA.Text = ""
End Sub
