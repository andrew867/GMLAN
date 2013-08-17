VERSION 5.00
Begin VB.Form dlg_Log 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Running Log"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_Log 
      Height          =   3255
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Dialog.frx":0000
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "dlg_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Start Logging
txt_Log.SelStart = Len(txt_Log)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Stop Logging
End Sub

