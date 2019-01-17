---
title: FileDialog.Show method (Office)
keywords: vbaof11.chm256012
f1_keywords:
- vbaof11.chm256012
ms.prod: office
api_name:
- Office.FileDialog.Show
ms.assetid: e67f7fc3-326d-12d0-fe44-e20048ff6abf
ms.date: 01/09/2019
localization_priority: Priority
---


# FileDialog.Show method (Office)

Displays a file dialog box and returns a **Long** indicating whether the user pressed the **Action** button (-1) or the **Cancel** button (0). When you call the **Show** method, no more code executes until the user dismisses the file dialog box. In the case of the **Open** and **SaveAs** dialog boxes, use the **Execute** method right after the **Show** method to carry out the user's action.


## Syntax

_expression_.**Show**

_expression_ Required. A variable that represents a **[FileDialog](Office.FileDialog.md)** object.


## Example

The following example displays a **File Picker** dialog box by using the **FileDialog** object, and displays each selected file in a message box.


```vb
Sub Main() 
 
 'Declare a variable as a FileDialog object. 
 Dim fd As FileDialog 
 
 'Create a FileDialog object as a File Picker dialog box. 
 Set fd = Application.FileDialog(msoFileDialogFilePicker) 
 
 'Declare a variable to contain the path 
 'of each selected item. Even though the path is aString, 
 'the variable must be a Variant because For Each...Next 
 'routines only work with Variants and Objects. 
 Dim vrtSelectedItem As Variant 
 
 'Use a With...End With block to reference the FileDialog object. 
 With fd 
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'The user pressed the button. 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection. 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is a string that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "The path is: " &amp; vrtSelectedItem 
 
 Next vrtSelectedItem 
 'The user pressed Cancel. 
 Else 
 End If 
 End With 
 
 'Set the object variable to nothing. 
 Set fd = Nothing 
 
End Sub
```


## See also

- [FileDialog object members](overview/library-reference/filedialog-members-office.md)

