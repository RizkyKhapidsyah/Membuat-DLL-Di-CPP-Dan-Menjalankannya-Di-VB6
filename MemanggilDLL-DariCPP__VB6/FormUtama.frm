VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Membaca File DLL"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6615
      Begin VB.TextBox textInput1 
         Height          =   405
         Left            =   1800
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox textInput2 
         Height          =   405
         Left            =   1800
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   4695
      End
      Begin VB.CommandButton cmdHitung 
         Caption         =   "Command1"
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox textHasil 
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton cmdKosongkan 
         Caption         =   "Command1"
         Height          =   495
         Left            =   3240
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Input 1"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Input 2"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasil"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   345
      End
   End
   Begin VB.Label LabelJudul 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHitung_Click()
On Error GoTo HancurkanError
    If textInput1.Text = "" Then
        MsgBox "Input1 Kosong", vbExclamation + vbOKOnly, "Input Kosong"
        textInput1.SetFocus
    ElseIf textInput2.Text = "" Then
        MsgBox "Input2 Kosong", vbExclamation + vbOKOnly, "Input Kosong"
        textInput2.SetFocus
    Else
        textHasil.Text = Hitung(Val(textInput1.Text), Val(textInput2.Text))
        textHasil.SetFocus
    End If
Exit Sub
HancurkanError:
    If Err.Number = 6 Then
        MsgBox "Input terlalu banyak!", vbExclamation + vbOKOnly, "Error"
        textInput1.SetFocus
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdKosongkan_Click()
    With Me
        .textInput1.Text = ""
        .textInput2.Text = ""
        .textHasil.Text = ""
        .textInput1.SetFocus
    End With
End Sub

'Created by Rizky Khapidsyah

Private Sub Form_Load()
    With Me
        .LabelJudul.Caption = "Program ini membaca file DLL yang dibuat di Visual C++ yang berisi fungsi untuk perkalian :"
        .Frame1.Caption = Empty
        .textInput1.Text = ""
        .textInput2.Text = ""
        .textHasil.Text = ""
        .textHasil.Locked = True
        .cmdHitung.Caption = "Hitung"
        .cmdHitung.Font.Bold = True
        .cmdKosongkan.Caption = "Kosongkan"
    End With
End Sub
