---
title: FileDialog.DialogType property (Office)
keywords: vbaof11.chm256010
f1_keywords:
- vbaof11.chm256010
ms.prod: office
api_name:
- Office.FileDialog.DialogType
ms.assetid: c589fe49-6527-7cdc-b7cb-55ac71013f3c
ms.date: 01/09/2019
localization_priority: Normal
---


# FileDialog.DialogType property (Office)

Gets an **[MsoFileDialogType](office.msofiledialogtype.md)** constant representing the type of file dialog box that the **FileDialog** object is set to display. Read-only.


## Syntax

_expression_.**DialogType**

_expression_ A variable that represents a **[FileDialog](Office.FileDialog.md)** object.


## Example

The following example takes a **FileDialog** object of an unknown type and runs the **Execute** method if it is a **SaveAs** dialog box or an **Open** dialog box.


```vb
Sub DisplayAndExecuteFileDialog(ByRef fd As FileDialog) 
 
 'Use a With...End With block to reference the FileDialog object. 
 With fd 
 'If the user presses the action button... 
 If .Show = -1 Then 
 
 'Use the DialogType property to determine whether to 
 'use the Execute method. 
 Select Case .DialogType 
 Case msoFileDialogOpen, msoFileDialogSaveAs: .Execute 
 'Do nothing otherwise. 
 Case Else 
 End Select 
 'If the user presses Cancel... 
 Else 
 End If 
 End With 
 
End Sub
```


## See also

- [FileDialog object members](overview/library-reference/filedialog-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]