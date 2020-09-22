VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColorize 
   Caption         =   "Colorize VB Code"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Paste VB Code from Clipboard"
      Height          =   555
      Left            =   2760
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9763
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   99999
      TextRTF         =   $"frmColorize.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Colorize rtb, Clipboard.GetText
End Sub
