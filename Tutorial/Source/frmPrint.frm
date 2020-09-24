VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aeric Cake House"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3225
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Catalogue"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
frmReport.WindowState = 2
frmReport.Show
Unload Me
End Sub

Private Sub Form_Load()
If OpenDatabase = False Then
    Unload Me
    Exit Sub
End If
End Sub
