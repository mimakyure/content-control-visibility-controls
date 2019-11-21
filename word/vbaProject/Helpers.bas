Attribute VB_Name = "Helpers"
'===================================================================================================
' Set of generally useful procedures

' Test whether a key is present in a collection
Function Exists(Coll As Collection, Key As String) As Boolean

  On Error Resume Next
  Coll.Item Key
  Exists = (Err.Number = 0)
  Err.Clear

End Function

' Create collection from two arrays
Function CreateCollection(Items() As Variant, Keys() As Variant) As Collection

  Dim Idx As Integer
  Dim Coll As New Collection
  
  ' Store the key with the data
  For Idx = LBound(Keys) To UBound(Keys)
    Coll.Add Array(Items(Idx), Keys(Idx)), Keys(Idx)
  Next Idx
  
  Set CreateCollection = Coll

End Function

' Retrieve item from collection if it exists or return Empty
Function GetCollectionItem(Coll As Collection, Key As String) As Variant

  Dim Item As Variant
  
  If Exists(Coll, Key) Then
    Item = Coll.Item(Key)
  Else
    Item = Empty
  End If
  
  GetCollectionItem = Item
    
End Function

' Test if value exists and matches target value
Function TestVar(Val As Variant, TestVal As Variant) As Boolean

  Dim EmptyTest As Boolean
  Dim ValTest As Boolean
  
  EmptyTest = IsEmpty(Val)
  ValTest = TestVal = Val
  
  If Not EmptyTest Then ValTest = (TestVal = Val)
  
  TestVar = (EmptyTest Or ValTest)
  
End Function

' Clear any previously set find paramaters
Sub ResetRangeFind(Rng As Range)

  With Rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = ""
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
  End With
  
End Sub

' Prepare to make changes that will alter the view
Function PrepareForChanges() As Collection

  Dim Settings As New Collection
  
  ' Prevent intermediate updating
  Settings.Add Application.ScreenUpdating, "ScreenUpdating"
  Application.ScreenUpdating = False
  
  ' Show formatting so it can be adjusted
  With ActiveDocument.ActiveWindow.View
    Settings.Add .ShowAll, "ShowAll"
    Settings.Add .ShowHiddenText, "ShowHiddenText"
    .ShowAll = True
    .ShowHiddenText = False
  End With

  Set PrepareForChanges = Settings
  
End Function


Sub FinishChanges(Settings As Collection)

  ActiveDocument.ActiveWindow.View.ShowAll = Settings("ShowAll")
  ActiveDocument.ActiveWindow.View.ShowHiddenText = Settings("ShowHiddenText")
  Application.ScreenUpdating = Settings("ScreenUpdating")
  
End Sub
