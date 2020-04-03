---
title: Application.FileDialog property (PowerPoint)
keywords: vbapp10.chm502046
f1_keywords:
- vbapp10.chm502046
ms.prod: powerpoint
api_name:
- PowerPoint.Application.FileDialog
ms.assetid: 0f0d5b6c-e478-6d15-7218-be04df978d6b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FileDialog property (PowerPoint)

Returns a  **FileDialog** object that represents a single instance of a file dialog box. Read-only.


## Syntax

_expression_. `FileDialog`( `_Type_` )

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoFileDialogType**|The type of dialog to return.|

## Return value

FileDialog


## Remarks

The value of the Type parameter can be one of these  **MsoFileDialogType** constants.


||
|:-----|
|**msoFileDialogFilePicker**|
|**msoFileDialogFolderPicker**|
|**msoFileDialogOpen**|
|**msoFileDialogSaveAs**|

## Example

This example displays the  **Save As** dialog box.


```vb
Sub ShowSaveAsDialog()

    Dim dlgSaveAs As FileDialog

    Set dlgSaveAs = Application.FileDialog( _
        Type:=msoFileDialogSaveAs)

    dlgSaveAs.Show

End Sub
```

This example displays the  **Open** dialog box and allows a user to select multiple files to open.




```vb
Sub ShowFileDialog()

    Dim dlgOpen As FileDialog

    Set dlgOpen = Application.FileDialog( _
        Type:=msoFileDialogOpen)

    With dlgOpen
        .AllowMultiSelect = True
        .Show
    End With

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]