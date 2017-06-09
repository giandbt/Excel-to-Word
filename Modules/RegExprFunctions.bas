Attribute VB_Name = "RegExprFunctions"
'Source:  Patrick Matthews @ E-E Article:  http://www.experts-exchange.com/Programming/Languages/Visual_Basic/A_1336-Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html?sfQueryTermInfo=1+30+matthewspatrick+regular
'Modifications by DLMILLE @E-E on 11-5-2011 to return Null as opposed to "" with the RegExpFind function, to test for IsNull on expected array result
' Module level declarations
 
Option Explicit
Option Compare Text
 
Function RegExpFind(LookIn As String, PatternStr As String, Optional Pos, _
    Optional MatchCase As Boolean = True, Optional ReturnType As Long = 0, _
    Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string (LookIn), and return matches to a
    ' pattern (PatternStr).  Use Pos to indicate which match you want:
    ' Pos omitted               : function returns a zero-based array of all matches
    ' Pos = 1                   : the first match
    ' Pos = 2                   : the second match
    ' Pos = <positive integer>  : the Nth match
    ' Pos = 0                   : the last match
    ' Pos = -1                  : the last match
    ' Pos = -2                  : the 2nd to last match
    ' Pos = <negative integer>  : the Nth to last match
    ' If Pos is non-numeric, or if the absolute value of Pos is greater than the number of
    ' matches, the function returns an empty string.  If no match is found, the function returns
    ' an empty string.  (Earlier versions of this code used zero for the last match; this is
    ' retained for backward compatibility)
    
    ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
    ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
    
    ' ReturnType indicates what information you want to return:
    ' ReturnType = 0            : the matched values
    ' ReturnType = 1            : the starting character positions for the matched values
    ' ReturnType = 2            : the lengths of the matched values
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function in Excel, you can use range references for any of the arguments.
    ' If you use this in Excel and return the full array, make sure to set up the formula as an
    ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
    
    ' Note: RegExp counts the character positions for the Match.FirstIndex property as starting
    ' at zero.  Since VB6 and VBA has strings starting at position 1, I have added one to make
    ' the character positions conform to VBA/VB6 expectations
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim Answer()
    Dim Counter As Long
    
    ' Evaluate Pos.  If it is there, it must be numeric and converted to Long
    
    If Not IsMissing(Pos) Then
        If Not IsNumeric(Pos) Then
            RegExpFind = Null 'updated by dlmille of E-E for testing of null result if empty or array without generating error
            Exit Function
        Else
            Pos = CLng(Pos)
        End If
    End If
    
    ' Evaluate ReturnType
    
    If ReturnType < 0 Or ReturnType > 2 Then
        RegExpFind = Null 'updated by dlmille of E-E for testing of null result if empty or array without generating error
        Exit Function
    End If
    
    ' Create instance of RegExp object if needed, and set properties
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
        
    ' Test to see if there are any matches
    
    If RegX.test(LookIn) Then
        
        ' Run RegExp to get the matches, which are returned as a zero-based collection
        
        Set TheMatches = RegX.Execute(LookIn)
        
        ' Test to see if Pos is negative, which indicates the user wants the Nth to last
        ' match.  If it is, then based on the number of matches convert Pos to a positive
        ' number, or zero for the last match
        
        If Not IsMissing(Pos) Then
            If Pos < 0 Then
                If Pos = -1 Then
                    Pos = 0
                Else
                    
                    ' If Abs(Pos) > number of matches, then the Nth to last match does not
                    ' exist.  Return a zero-length string
                    
                    If Abs(Pos) <= TheMatches.Count Then
                        Pos = TheMatches.Count + Pos + 1
                    Else
                        RegExpFind = Null 'updated by dlmille of E-E for testing of null result if empty or array without generating error
                        GoTo Cleanup
                    End If
                End If
            End If
        End If
        
        ' If Pos is missing, user wants array of all matches.  Build it and assign it as the
        ' function's return value
        
        If IsMissing(Pos) Then
            ReDim Answer(0 To TheMatches.Count - 1)
            For Counter = 0 To UBound(Answer)
                Select Case ReturnType
                    Case 0: Answer(Counter) = TheMatches(Counter)
                    Case 1: Answer(Counter) = TheMatches(Counter).FirstIndex + 1
                    Case 2: Answer(Counter) = TheMatches(Counter).Length
                End Select
            Next
            RegExpFind = Answer
        
        ' User wanted the Nth match (or last match, if Pos = 0).  Get the Nth value, if possible
        
        Else
            Select Case Pos
                Case 0                          ' Last match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(TheMatches.Count - 1)
                        Case 1: RegExpFind = TheMatches(TheMatches.Count - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(TheMatches.Count - 1).Length
                    End Select
                Case 1 To TheMatches.Count      ' Nth match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(Pos - 1)
                        Case 1: RegExpFind = TheMatches(Pos - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(Pos - 1).Length
                    End Select
                Case Else                       ' Invalid item number
                    RegExpFind = Null 'updated by dlmille of E-E for testing of null result if empty or array without generating error
            End Select
        End If
    
    ' If there are no matches, return empty string
    
    Else
        RegExpFind = Null 'updated by dlmille of E-E for testing of null result if empty or array without generating error
    End If
    
Cleanup:
    ' Release object variables
    
    Set TheMatches = Nothing
    
End Function
