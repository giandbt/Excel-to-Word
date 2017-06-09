VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BM_Indicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' David L. Miller, Nov, 2011
' Created 2011-11-22

Option Explicit
Option Compare Binary

' Containers for "write-once read-many" property
Private Safe_Name As String
Private Safe_Obj As Object
Private Safe_Type As Integer

Property Get Name() As String
Attribute Name.VB_Description = "Key value used when adding BM_Indicator item to the BM_Indicators parent collection.  Once written, it cannot be changed"
    
    ' Returns item's Name value
    
    Name = Safe_Name
    
End Property

Property Let Name(NameString As String)
    
    ' Sets Name value for item
    
    ' This makes the Name property "write once, read many".  If the Name is a zero length string,
    ' the Property Let allows you to change it; if not, the procedure raises a user defined
    ' error.  Basically, we cannot allow changes because we want this property to match the
    ' item's true key used when it was added to the parent BM_Indicators collection
    
    If Safe_Name = "" Then
        Safe_Name = NameString
    Else
        Err.Raise vbObjectError + 1001, , "Cannot change Name property of BM_Indicator object"
    End If
    
End Property
Property Get obj() As Object
        
    Set obj = Safe_Obj
    
End Property

Property Let obj(objObject As Object)
    
    If Safe_Obj Is Nothing Then
        Set Safe_Obj = objObject
    Else
        Err.Raise vbObjectError + 1001, , "Cannot change obj property of BM_Indicator object"
    End If
    
End Property
Property Get BM_Type() As Integer
    
    ' Returns item's Name value
    
    BM_Type = Safe_Type
    
End Property

Property Let BM_Type(IntType As Integer)
    
    If Safe_Type = 0 Then
        Safe_Type = IntType
    Else
        Err.Raise vbObjectError + 1001, , "Cannot change Name property of BM_Indicator object"
    End If
    
End Property
