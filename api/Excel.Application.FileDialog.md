---
title: Application.FileDialog property (Excel)
keywords: vbaxl10.chm133270
f1_keywords:
- vbaxl10.chm133270
ms.prod: excel
api_name:
- Excel.Application.FileDialog
ms.assetid: 96a6fdc5-1bde-68dd-2493-9d8a92915afb
ms.date: 04/04/2019
localization_priority: Priority
---


# Application.FileDialog property (Excel)

Returns a **[FileDialog](Office.FileDialog.md)** object representing an instance of the file dialog.


## Syntax

_expression_.**FileDialog** (_fileDialogType_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _fileDialogType_|Required| **[MsoFileDialogType](Office.MsoFileDialogType.md)**|The type of file dialog.|

## Remarks

**MsoFileDialogType** can be one of these constants:

- **msoFileDialogFilePicker**. Allows user to select a file.
- **msoFileDialogFolderPicker**. Allows user to select a folder.
- **msoFileDialogOpen**. Allows user to open a file.
- **msoFileDialogSaveAs**. Allows user to save a file.

## Example

In this example, Microsoft Excel opens the file dialog allowing the user to select one or more files. After these files are selected, Excel displays the path for each file in a separate message.

```vb
Sub UseFileDialogOpen() 
 
    Dim lngCount As Long 
 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
 
    End With 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
