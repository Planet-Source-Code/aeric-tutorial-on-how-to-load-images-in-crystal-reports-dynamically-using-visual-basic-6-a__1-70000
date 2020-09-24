VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReport 
   Caption         =   "Catalogue"
   ClientHeight    =   4575
   ClientLeft      =   5175
   ClientTop       =   4170
   ClientWidth     =   7560
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   7560
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rpt As CRAXDRT.Report
Dim db As CRAXDRT.Database
Dim rs As New ADODB.Recordset
Dim WithEvents sect As CRAXDRT.Section
Attribute sect.VB_VarHelpID = -1

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Set rpt = crx.OpenReport(App.Path & "\Report\Catalogue.rpt")
Set db = rpt.Database
Set sect = rpt.Sections("Section5")
rs.Open "SELECT * FROM Cake", cn, 1, 1
rpt.Database.SetDataSource rs, 3, 1
CRViewer1.ReportSource = rpt
CRViewer1.ViewReport
CRViewer1.Zoom 1
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub

Private Sub sect_Format(ByVal pFormattingInfo As Object)
Dim bmp As StdPicture
On Error Resume Next
With sect.ReportObjects
    Set .Item("picCake").FormattedPicture = LoadPicture(App.Path & "\Cake\ck99.gif") 'default
    If .Item("adoFileName").Value <> "" Then
        Set bmp = LoadPicture(App.Path & "\Cake\" & .Item("adoFileName").Value)
        Set .Item("picCake").FormattedPicture = bmp
    End If
End With

End Sub
