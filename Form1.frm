VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin BlackBeautyButton.BlackBeauty BlackBeauty4 
      Height          =   435
      Left            =   135
      TabIndex        =   3
      Top             =   1755
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      Caption         =   "Done!!"
      Forecolor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Balloon"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverCaptionColor=   16777215
   End
   Begin BlackBeautyButton.BlackBeauty BlackBeauty3 
      Height          =   435
      Left            =   135
      TabIndex        =   2
      Top             =   1170
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      Caption         =   "yet another"
      Forecolor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MouseOverCaptionColor=   16776960
   End
   Begin BlackBeautyButton.BlackBeauty BlackBeauty2 
      Height          =   435
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      Caption         =   "Nother button"
      Forecolor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BlackBeautyButton.BlackBeauty BlackBeauty1 
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      Caption         =   "Button1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverCaptionColor=   49344
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BlackBeauty1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Debug.Print "mouse down"
End Sub

Private Sub BlackBeauty1_MouseEnter()
   Debug.Print "mouse enter"
End Sub

Private Sub BlackBeauty1_MouseExit()
  Debug.Print "mouse exit"
End Sub

Private Sub BlackBeauty1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Debug.Print "mouse up"
End Sub
