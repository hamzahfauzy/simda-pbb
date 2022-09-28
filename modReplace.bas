Attribute VB_Name = "modReplace"
Public Function Rep(ByVal Kata As String) As String
    Rep = Replace(Kata, "'", "`")
End Function

