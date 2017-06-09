VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BM_Indicators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' David L. Miller, Nov, 2011
' Created 2011-11-22

' Parent collection class of BM_Indicator

Option Explicit
Option Compare Binary

' Container for all BM_Indicator objects in the parent collection class
Private coll As Collection

Private Sub Class_Initialize()
    
    Set coll = New Collection
    
End Sub

Private Sub Class_Terminate()
    
    Set coll = Nothing
    
End Sub

Public Function Add(Name As String, obj As Object, BM_Type As Integer) As BM_Indicator
Attribute Add.VB_Description = "Adds new BM_Indicator item to the BM_Indicators parent collection"
    
    ' Adds a new item to the collection.  Causes an error if an item with the same key already exists
    ' or if you pass a zero length string for the Name argument
    
    If Name = "" Then
        Err.Raise vbObjectError + 1002, , "Name property of BM_Indicator object cannot be zero length string"
    End If
    
    Set Add = New BM_Indicator
    Add.Name = Name
    Add.obj = obj
    Add.BM_Type = BM_Type
    coll.Add Add, Name
    
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Recreates BM_Indicators instance, thus removing all members"
    
    ' Recreates (and thus clears) collection
    
    Set coll = New Collection
    
End Sub

Property Get Count() As Long
Attribute Count.VB_Description = "Returns count of members in BM_Indicators instance"
    
    ' Returns number of items in the collection
    
    ' Read-only
    
    Count = coll.Count
    
End Property

Property Get Item(Index As Variant) As BM_Indicator
Attribute Item.VB_Description = "Returns a BM_Indicator member from BM_Indicators parent collection"
Attribute Item.VB_UserMemId = 0
    
    ' Default property.  Returns an item from the collection.  Index may be either ordinal position (Long) or Name (String)
    
    ' Read-only
    
    On Error Resume Next
    Set Item = coll(Index)
    On Error GoTo 0
    
End Property

Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a BM_Indicator member from BM_Indicators instance"
    
    ' Removes an item from the collection.  Index may be either ordinal position (Long) or Name (String)
    
    coll.Remove Index
    
End Sub

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    ' Enables enumeration of the BM_Indicators parent collection, i.e.:
    '
    ' For Each Child In Parent...Next
    
    Set NewEnum = coll.[_NewEnum]
End Function

Function RemoveAll()
Dim i As Long

    If Not coll Is Nothing Then
        For i = 1 To coll.Count
            coll.Remove 1
        Next i
        Set coll = Nothing
    End If
    
End Function
