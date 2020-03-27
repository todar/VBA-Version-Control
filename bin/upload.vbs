' This opens up the main Excel app and call's the code
' to import all components from the src directory.
' This code should be moved into this script in the future,
' but for testing it is easier to write in Excel.
Option Explicit

' The app Filepath to the app will be passed in through the cli script.
Dim filepath
filepath = WScript.Arguments(0)

' This might be dangerous depending on how I implement the final code.
' For instance, this might remove all code from the project to ensure it
' is kept clean. Please see readme for more details.
Sub dangerouslyImportComponents() 
  ' Need an instance of Excel in order to open the workbook.
  Dim app 
  Set app = CreateObject("Excel.Application") 

  ' Open the app workbook and call the function to import from
  ' the src folder.
  Dim workbookRef
  Set workbookRef = app.Workbooks.Open(filepath, 0, True) 
  app.Run ""
  app.Quit 

  Set workbookRef = Nothing 
  Set app = Nothing
End Sub 

' Call to the main function.
dangerouslyImportComponents