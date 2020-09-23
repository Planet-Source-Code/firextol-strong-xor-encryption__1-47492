VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Strong Encryption"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDc 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Decrypted..."
      Top             =   1560
      Width           =   6255
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPw 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Password"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtEn 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   6255
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtIn 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "To encypt"
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function AsctoHex(ByVal astr As String)
For x = 1 To Len(astr)
hc = Hex$(Asc(Mid$(astr, x, 1)))
nstr = nstr & String(2 - Len(hc), "0") & hc
Next
AsctoHex = nstr
End Function

Private Function HexToAsc(ByVal hstr As String)
For x = 1 To Len(hstr) Step 2
nstr = nstr & Chr(Val("&H" & Mid$(hstr, x, 2)))
Next
HexToAsc = nstr
End Function

Private Sub cmdDecrypt_Click()
Dim x As Long
Dim eKey As Long, eChr As Byte, oChr As Byte, tmp$
    
Rnd -1
Randomize Len(txtPw.Text)

For i = 1 To Len(txtPw.Text)
    'generate a key based on pw
    eKey = eKey + (Asc(Mid$(txtPw.Text, i, 1)) Xor Fix(255 * Rnd) Xor (i Mod 256))
Next
'reset random function
Rnd -1
'initilize our key as the random seed
Randomize eKey
'generate a pseudo old char, which makes decryption dependent on the preceding encrypted character
oChr = Int(Rnd * 256)
'start decryption
tmp$ = HexToAsc(txtEn.Text)
txtDc.Text = ""
For x = 1 To Len(tmp$)
    pp = pp + 1
    If pp > Len(txtPw.Text) Then pp = 1
    If x > 1 Then oChr = Asc(Mid$(tmp$, x - 1, 1))
    eChr = Asc(Mid$(tmp$, x, 1)) Xor Int(Rnd * 256) Xor Asc(Mid$(txtPw.Text, pp, 1)) Xor oChr
    txtDc.Text = txtDc.Text & Chr$(eChr)
Next
End Sub

Private Sub cmdEncrypt_Click()
Dim x As Long
Dim eKey As Long, eChr As Byte, oChr As Byte, tmp$

Rnd -1
Randomize Len(txtPw.Text)

For i = 1 To Len(txtPw.Text)
    'generate a key based on pw
    eKey = eKey + (Asc(Mid$(txtPw.Text, i, 1)) Xor Fix(255 * Rnd) Xor (i Mod 256))
Next
'reset random function
Rnd -1
'initilize our key as the random seed
Randomize eKey
'generate a pseudo old char, which makes decryption dependent on other characters, teehee!
oChr = Int(Rnd * 256)
'start encryption
For x = 1 To Len(txtIn.Text)
    pp = pp + 1
    If pp > Len(txtPw.Text) Then pp = 1
    eChr = Asc(Mid$(txtIn.Text, x, 1)) Xor Int(Rnd * 256) Xor Asc(Mid$(txtPw.Text, pp, 1)) Xor oChr
    tmp$ = tmp$ & Chr(eChr)
    oChr = eChr
Next
txtEn.Text = AsctoHex(tmp$)
End Sub

