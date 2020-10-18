Attribute VB_Name = "mSystemFunctions"
Option Explicit

Public Sub Freeze()

    'Switch off screen updating and display alerts
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False

End Sub

Public Sub Thaw()

    'Switch on screen updating and display alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Public Sub loadRcLoader()
    mControl.loadRCsToCashbox
End Sub
