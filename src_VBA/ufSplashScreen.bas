Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:03"), "CloseSplashScreen"
End Sub



