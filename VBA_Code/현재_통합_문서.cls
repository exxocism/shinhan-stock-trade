VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Module1.QuitApp
End Sub

Private Sub Workbook_Open()

    Sheet1.MakeControl
    Sheet1.Load_settings
    Sheet1.CheckBox4.Value = Sheets("v").Range("USE_PLUGIN").Value2
    
End Sub
