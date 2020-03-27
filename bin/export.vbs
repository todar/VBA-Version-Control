' This opens up the main Excel app and call's the code
' to export all components into the src directory.
' This code should be moved into this script in the future,
' but for testing it is easier to write in Excel.
Option Explicit

' The app Filepath to the app will be passed in through the cli script.
Dim filepath
filepath = WScript.Arguments(0)

Sub exportComponents() 
  ' Need an instance of Excel in order to open the workbook.
  Dim app 
  Set app = CreateObject("Excel.Application") 

  ' Open the app workbook and call the function to export.
  Dim workbookRef
  Set workbookRef = app.Workbooks.Open(filepath, 0, True) 
  app.Run "exportComponents"
  app.Quit 

  Set workbookRef = Nothing 
  Set app = Nothing
End Sub 

' Call to the main function.
exportComponents