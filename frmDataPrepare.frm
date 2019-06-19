VERSION 5.00
Begin VB.Form frmDataPrepare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2385
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   2055
      Width           =   1530
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   435
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2055
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1350
      Width           =   1560
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   510
      Left            =   240
      TabIndex        =   0
      Top             =   495
      Width           =   1575
   End
End
Attribute VB_Name = "frmDataPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

