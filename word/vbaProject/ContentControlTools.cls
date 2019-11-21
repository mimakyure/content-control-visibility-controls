VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentControlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================
' Content control visibility/text hidden attribute modifiers

Public DisableWindowSelectionChange As Boolean

' Set font hidden property on style separator paragraph mark
' This is used to keep style separators hidden when unhiding a range
Private Sub HideStyleSeparators(Rng As Range)
    
  Dim Pars As Variant
  Dim Par As Paragraph
  
  For Each Par In Rng.Paragraphs
  
    ' Hidden needed to detect style separator
    Par.Range.Font.Hidden = True
    
    ' Unhide everthing but the style separator at the end
    If (Par.IsStyleSeparator) Then
      With Par.Range
        .MoveEnd wdCharacter, -1
        .Font.Hidden = False
      End With
    Else
      Par.Range.Font.Hidden = False
    End If
    
  Next Par
  
End Sub

' Get range containing first non-hidden character from start of line
Private Function GetFirstNotHidden(CC As ContentControl) As Range

  Dim Rng As Range: Set Rng = CC.Range
  Rng.Collapse wdCollapseStart

  ' Go to beginning of line
  ResetRangeFind Rng
  Rng.Find.Forward = False
  Rng.Find.Font.Hidden = False
  Rng.Find.MatchWildcards = True
  Rng.Find.Text = "[^13^11]"
  Rng.Find.Execute
  If Rng.start = Rng.End Then
    Rng.start = ActiveDocument.Range.start
    Rng.End = ActiveDocument.Range.start
  End If

  ' Find first non-hidden text
  Rng.Find.Forward = True
  Rng.Find.Text = ""
  Rng.Find.Execute

  Set GetFirstNotHidden = Rng
  
End Function


' Set hidden property of content control font
Private Sub HideContentControl(CC As ContentControl, Hidden As Boolean)

  Dim Rng As Range: Set Rng = CC.Range
  Dim VisibleRng As Range
  
  ' Include content control end blocks
  Rng.MoveStart wdCharacter, -1
  Rng.MoveEnd wdCharacter, 1
  Rng.Font.Hidden = Hidden
  
  ' Hide line if all text is hidden
  If (Hidden) Then
  
    Set VisibleRng = GetFirstNotHidden(CC)
    If VisibleRng.Text = Chr(13) Or _
       VisibleRng.Text = Chr(11) Then
       VisibleRng.Font.Hidden = True
    End If
  
  ' Keep style separators hidden when unhiding
  Else
    HideStyleSeparators Rng
  End If
  
End Sub

' Set font hidden property on content controls with matching title, tag, and category
Sub HideControls(Hidden As Boolean, _
                 Optional Title As Variant = Empty, Optional Tag As Variant = Empty, _
                 Optional Gallery As Variant = Empty, Optional Category As Variant = Empty)

  Dim Settings As Collection: Set Settings = PrepareForChanges
  
  Dim CC As ContentControl
  Dim Rng As Range
   
  ' Set visibility of matching controls
  For Each CC In ActiveDocument.ContentControls

    If (TestVar(Title, CC.Title) And TestVar(Tag, CC.Tag) And _
        TestVar(Gallery, CC.BuildingBlockType) And _
        TestVar(Category, CC.BuildingBlockCategory)) Then
   
      HideContentControl CC, Hidden
      
    End If
  Next CC
  
  ' Keep visible portion of document about the same
  ActiveWindow.ScrollIntoView Selection.Range, True
  FinishChanges Settings
  
End Sub

' Identify selected content controls and set hidden property
Sub SetSelectionHidden(Hidden As Boolean)

  Dim Settings As Collection: Set Settings = PrepareForChanges
  
  Dim CC As ContentControl
  Dim Rng As Range
  
  ' Hide content controls in selection
  If (Selection.Range.ContentControls.Count > 0) Then
    For Each CC In Selection.Range.ContentControls
      HideContentControl CC, Hidden
    Next CC
  End If
  
  ' Hide content controls at cursor location or partially selected
  If (Not Selection.Range.ParentContentControl Is Nothing) Then
    HideContentControl Selection.Range.ParentContentControl, Hidden
  End If
  
  FinishChanges Settings
  
End Sub

'===================================================================================================
' Content control appearance

' Set appearance of all content controls
Public Sub ControlAppearance(Appearance As WdContentControlAppearance)

  Dim Settings As Collection: Set Settings = PrepareForChanges
  Dim CC As ContentControl
  
  For Each CC In ActiveDocument.ContentControls
    CC.Appearance = Appearance
  Next CC
  
  ' Keep visible portion of document about the same
  ActiveWindow.ScrollIntoView Selection.Range, True
  FinishChanges Settings
  
End Sub

'====================================================================================================
' Ribbon setup

' Get list of all content control values with matching property name
Function GetPropertyValues(Properties As Collection) As Collection

  Dim Prop As Variant
  Dim PropVals As New Collection
  Dim CC As ContentControl
  Dim Val As Variant
  Dim Types As BuildingBlockTypes
  Dim GalleryNames As New Collection
  
  ' Set up collections to store values
  For Each Prop In Properties
    PropVals.Add New Collection, Prop(1)
  Next Prop
  
  ' Lookup property values
  On Error Resume Next
  For Each CC In ActiveDocument.ContentControls
    For Each Prop In Properties
      Val = CallByName(CC, Prop(0), VbGet)
      PropVals.Item(Prop(1)).Add Val, CStr(Val)
    Next Prop
  Next CC

  ' Retrieve gallery names
  If Exists(PropVals, "Gallery") Then
    Set Types = ActiveDocument.AttachedTemplate.BuildingBlockTypes
    For Each Val In PropVals.Item("Gallery")
      GalleryNames.Add Types(Val).Name, Types(Val).Name
    Next Val
    PropVals.Remove "Gallery"
    PropVals.Add GalleryNames, "Gallery"
  End If
  
  Set GetPropertyValues = PropVals

End Function


