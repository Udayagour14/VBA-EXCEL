# VBA-EXCEL
Visual Basic for Applications (VBA) is a programming language used to automate tasks and enhance the functionality of Microsoft Office applications. You can use it with Excel, Word, Access, and other Office programs, althought it’s probably most widely used for creating custom macros in Excel
### list of some of the most used terms you'll encounter as you start automating tasks and building custom solutions in Excel.
Modules are the containers for VBA code, where procedures and functions are stored.
Objects are the building blocks of VBA. They represent elements like workbooks, worksheets, and cells.
Procedures are the blocks of code that perform specific tasks, often categorized as sub-procedures or functions.
Statements are the instructions within a procedure that tell Excel (or Word or Access) what actions to perform.
Variables store data that can be used and manipulated within your code.
Logical operators compare values and make decisions based on the results. They include operators like And, Or, and Not.

#### The Workbook object is any Excel file that’s currently open. It allows us to perform actions such as adding new sheets or saving the existing sheets within the workbook. In this example, I want to add a sheet to the currently opened sheet and then save it. At first, I had Sheet1. To add another sheet, I write the following code:
 Sub AddSheetAndSaveWorkbook()
    ' Adds a new worksheet to the active workbook
   ActiveWorkbook.Sheets.Add
    ActiveWorkbook.Save
End Sub

![image](https://github.com/user-attachments/assets/26c6223e-8ab2-4383-9b5d-45e112022b66)

#### The Worksheet object represents the currently active sheet in Excel. With this, you can modify or manipulate the active sheet. For example, I want to change the name of the active sheet. To do so, I enter the following code:
Sub RenameActiveSheet()
    ' Renames the active sheet to "Sales Report".
    ActiveSheet.Name = "Sales Report"
End Sub

![image](https://github.com/user-attachments/assets/cf9561ea-90b6-4f92-b768-0120d9bad111)

Automating Tasks with VBA Excel Macros
Macros are essentially a series of instructions created in VBA to perform repetitive tasks. Using a macro, you can record a series of actions, like formatting cells, copying data, or performing calculations. After saving the macro, we can re-apply these actions with a single click. This saves time, especially when working with large datasets or tasks.

For example, if you often format your reports the same way, you can record a macro that applies all the necessary formatting steps instead of doing it manually each time. Later, you can run this macro to format new data. We showed a couple of basic examples, but imagine doing more small actions with just one click. Even if you don’t have coding experience, you can still automate routine tasks to streamline workflow and reduce the chances of errors that can occur when performing repetitive actions manually.

Recording an Excel macro 
Let’s see how you can record a macro to make some formatting changes:
Go to the Developer tab > Record Macro button. 
A dialog box will appear. Give your macro a name like I did FormatCells.
Assign a keyboard shortcut if desired. I assigned it Ctrl+S.
Click OK to start recording.

![image](https://github.com/user-attachments/assets/be87d93d-d30c-4d1d-998c-3a2495a3d629)



Now, if you have another dataset in another sheet and want the same formatting in that dataset, too, instead of formatting the whole thing again, press the shortcut key you created (in my case, it’s Ctrl+S). Otherwise, go to sheet2 > Developer tab > Macros > Run to do it manually.
![image](https://github.com/user-attachments/assets/a112cea4-4046-4079-abe4-1fad62d21b01)
### loop
Sub LoopThroughRange()
    Dim cell As Range
    For Each cell In Range("A1:A5")
        cell.Value = "Hello " & cell.Row
    Next cell
End Sub

Private Sub Constant_demo_Click()  
   'fruits is an array
   fruits = Array("apple", "orange", "cherries")
   Dim fruitnames As Variant
 
   'iterating using For each loop.
   For Each Item In fruits
      fruitnames = fruitnames & Item & Chr(10)
   Next
   
   MsgBox fruitnames
End Sub


