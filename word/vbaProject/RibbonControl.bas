Attribute VB_Name = "RibbonControl"
Public CustomRibbon As IRibbonUI
Private CCProperties As Collection
Private SelectedProperties As New Collection
Private EvtsAppClass As New EventsApplication
Private CCT As New ContentControlTools

' Initialize
Private Sub Onload(Ribbon As IRibbonUI)

  Set CustomRibbon = Ribbon
  RefreshCCProperties
  EvtsAppClass.Init Word.Application

End Sub

'===================================================================================================
' Building blocks group

' Insert building block content control
Private Sub AddBuildingBlockContentControl(Control As IRibbonControl)
  Selection.Range.ContentControls.Add wdContentControlBuildingBlockGallery
End Sub

' Verify content control can be inserted at Selection
Private Sub AddBuildingBlockContentControlEnabled(Control As IRibbonControl, ByRef Enabled)
  Enabled = True
  CustomRibbon.InvalidateControl "cc-add-bbcc"
End Sub

' Open building block save dialog
Private Sub SaveQuickPart(Control As IRibbonControl)
  Dialogs(wdDialogCreateAutoText).Show
End Sub

' Enable save quick part button if text is selected
Private Sub SaveQuickPartEnabled(Control As IRibbonControl, ByRef HasSelection)
  HasSelection = Selection.End - Selection.start > 0
  CustomRibbon.InvalidateControl "cc-save-qp"
End Sub

' Open content control properties dialog
Private Sub ShowProperties(Control As IRibbonControl)
  Dialogs(wdDialogContentControlProperties).Show
End Sub

' Verify content control located at Selection
Private Sub ShowPropertiesEnabled(Control As IRibbonControl, ByRef HasCC)
  HasCC = Selection.Information(wdInContentControl)
  CustomRibbon.InvalidateControl "show-properties"
End Sub

'===================================================================================================
' Visibility group

' Callback for DropDown GetItemCount
Private Sub GetItemCount(Control As IRibbonControl, ByRef Count)
  Count = CCProperties.Item(Control.Tag).Count
End Sub

' Callback for DropDown GetItemLabel
Private Sub GetItemLabel(Control As IRibbonControl, Index, ByRef Label)
  Label = CCProperties.Item(Control.Tag).Item(Index + 1)
End Sub

' Callback DropDown GetSelectedIndex
Private Sub GetSelectedItemIndex(Control As IRibbonControl, ByRef Index)
  Index = 0
End Sub

' Callback for DropDown onAction
Private Sub DropDownChange(Control As IRibbonControl, id As String, Index As Integer)

  ' Delete current saved state
  If Exists(SelectedProperties, Control.Tag) Then
    SelectedProperties.Remove Control.Tag
  End If

  ' Prepare for changing visibility
  If Index > 0 Then
     SelectedProperties.Add CCProperties.Item(Control.Tag).Item(Index + 1), Control.Tag
     CustomRibbon.InvalidateControl "cc-hide-matching"
     CustomRibbon.InvalidateControl "cc-show-matching"
  End If
  
End Sub

' Update stored list of content control properties
Private Sub RefreshCCProperties()
  
  Dim PropNames() As Variant
  Dim SaveNames() As Variant
  Dim Galleries As Collection
  Dim Types As BuildingBlockTypes
  Dim Item As Variant
  Dim GalleryNames As New Collection
  Dim Name As Variant
  
  PropNames = Array("Title", "Tag", "BuildingBlockType", "BuildingBlockCategory")
  SaveNames = Array("Title", "Tag", "Gallery", "Category")
  
  ' Get properties of content controls in document
  Set CCProperties = CCT.GetPropertyValues(CreateCollection(PropNames, SaveNames))
  
  ' Add default strings for use in dropdown menus
  For Each Name In SaveNames
    If CCProperties.Item(Name).Count = 0 Then
      CCProperties.Item(Name).Add "Choose " & Name & "..."
    Else
      CCProperties.Item(Name).Add "Choose " & Name & "...", Before:=1
    End If
  Next Name

End Sub

' Refresh dropdown lists with current content control properties
Private Sub RefreshPropertyLists(Control As IRibbonControl)

  RefreshCCProperties
  CustomRibbon.InvalidateControl "cc-title-select"
  CustomRibbon.InvalidateControl "cc-tag-select"
  CustomRibbon.InvalidateControl "cc-gallery-select"
  CustomRibbon.InvalidateControl "cc-category-select"

End Sub

' Apply current dropdown Selection to hide/show matching controls
Private Sub SetVisibilityOfMatching(Hidden As Boolean)

  Dim Title As Variant
  Dim Tag As Variant
  Dim Gallery As Variant
  Dim Category As Variant
  
  Title = GetCollectionItem(SelectedProperties, "Title")
  Tag = GetCollectionItem(SelectedProperties, "Tag")
  Gallery = GetCollectionItem(SelectedProperties, "Gallery")
  Category = GetCollectionItem(SelectedProperties, "Category")
  
  CCT.HideControls Hidden, Title, Tag, Gallery, Category
   
End Sub

' Hide content controls matching selected properties
Private Sub HideMatching(Control As IRibbonControl)
  SetVisibilityOfMatching True
End Sub

' Show content controls matching selected properties
Private Sub ShowMatching(Control As IRibbonControl)
  SetVisibilityOfMatching False
End Sub

' Enable buttons to hide/show content controls if dropdown criteria is chosen
Private Sub HideShowMatchingEnabled(Control As IRibbonControl, ByRef HasChoice)
  HasChoice = SelectedProperties.Count > 0
End Sub

' Hide selected content controls
Private Sub HideSelection(Control As IRibbonControl)
  CCT.SetSelectionHidden True
End Sub

' Show selected content controls
Private Sub ShowSelection(Control As IRibbonControl)
  CCT.SetSelectionHidden False
End Sub

' Enable buttons to hide/show content controls currently selected in the document
Private Sub HideShowSelectionEnabled(Control As IRibbonControl, ByRef HasCCSelection)
  HasCCSelection = Selection.Range.ContentControls.Count > 0 Or _
                   Not Selection.Range.ParentContentControl Is Nothing
End Sub

' Hide all content controls
Private Sub HideAll(Control As IRibbonControl)
  CCT.HideControls True
End Sub

' Show all content controls
Private Sub ShowAll(Control As IRibbonControl)
  CCT.HideControls False
End Sub

' Toggle visibility of formatting marks
Private Sub ToggleShowAll(Control As IRibbonControl, Pressed)
  ActiveWindow.ActivePane.View.ShowAll = Pressed
End Sub

' Determine visibility of formatting marks
Private Sub ShowAllPressed(Control As IRibbonControl, ByRef Pressed)
  Pressed = ActiveWindow.ActivePane.View.ShowAll
End Sub

' Toggle design mode state
Private Sub ToggleDesignMode(Control As IRibbonControl, Pressed)
  ActiveDocument.ToggleFormsDesign
End Sub

' Determine design mode state
Private Sub DesignModePressed(Control As IRibbonControl, ByRef Pressed)
  Pressed = CommandBars("Control Toolbox").Visible
End Sub

'===================================================================================================
' Apperance group

' Set hidden appearance for all content controls
Private Sub SetAppearanceHidden(Control As IRibbonControl)
  CCT.ControlAppearance wdContentControlHidden
End Sub

' Set bounding box appearance for all content controls
Private Sub SetAppearanceBoundingBox(Control As IRibbonControl)
  CCT.ControlAppearance wdContentControlBoundingBox
End Sub

' Set tags appearance for all content controls
Private Sub SetAppearanceTags(Control As IRibbonControl)
  CCT.ControlAppearance wdContentControlTags
End Sub
