This program has largely been inspired by Ron de Bruin page https://www.rondebruin.nl/win/s9/win002.htm

In excel do:
Developer / Code / Macro Security / 
  Enable all macros
  Trust access to the VBA project object model

When Opening a workbook for macro the AutoRecover is set To False. It prevents being bothered by recovery functionality.

md5 from files to edit are checked against a cache file that is stored in srcs folder. Files are then updated only if needed.

issue:
exported file trying to run ExportModules issue: "user-defined type not defined" error message on this line: Dim VBComp As VBIDE.VBComponent
fix:
  Actually that might only be needed when using the export directly from a vba module not from python win32com
  add a reference to "Microsoft Visual Basic For Applications Extensibility" (in the VBA window, select Tools/References and set the check box for this).
  I am not sure about that one so I disable it for now ("Microsoft Scripting Runtime")
  If it is needed to check and add references programmatically that is the path to follow:
  ```python
    # That might be needed VBIDE is "Microsoft Visual Basic For Applications Extensibility"
    # refNames=[chkRef.Name for chkRef in wb.VBProject.References]
    # if "VBIDE" not in refNames:
    #     print("do something")
    # wb.VBProject.References.AddFromFile()
  ```

There is no way to get win32com object class members from python, in order to know them only the microsoft documentation can provide them: https://docs.microsoft.com/en-us/office/vba/api/overview/
For instance here for application object. https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)

Issue:
File always open in readonly ask me to recover. It was also an issue with a thread that I implemented in order to focus a window.
Solution:
Go in File/Options/Save: find the path and delete everything in that folder. Close excel/word and open again. Check if you get autorecovery files again or not.

Sheets vba can be exported easily but not imported. If imported they are not imported as sheets but as class modules. So any code that is global regarding to a sheet should be implemented in a vba module if possible. That is why sheets are not imported.

To Diable Annoying beep on compile errors MsgBox do:
Win+R
mmsys.cpl
go to Sounds tab
in scrolling list select Exclamation and sounds and select None