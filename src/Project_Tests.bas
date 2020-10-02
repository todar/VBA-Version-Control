Attribute VB_Name = "Project_Tests"
''
' Tests for the Project Module.
'
' This testing is a little bit different than normal,
' as what is being tested is really importing and exporting
' code modules...
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Module} Project
''
Option Explicit

' Demo for importing & exporting.
' Try changing this function and rerun your Export and Import to see if Git works.
Private Function add(ByVal a As Double, ByVal b As Double) As Double
    add = a + b
End Function

