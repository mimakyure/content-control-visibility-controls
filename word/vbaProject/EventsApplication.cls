VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1


' Enable/disable ribbon tab controls when selection changes
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)

  CustomRibbon.InvalidateControl "cc-save-qp"
  CustomRibbon.InvalidateControl "cc-add-bbcc"
  CustomRibbon.InvalidateControl "cc-show-properties"
  CustomRibbon.InvalidateControl "cc-hide-selection"
  CustomRibbon.InvalidateControl "cc-show-selection"
  
End Sub

' Bring in objects for event handling and managing ribbon control refreshing
Sub Init(ActiveApp As Application)
  Set App = ActiveApp
End Sub
