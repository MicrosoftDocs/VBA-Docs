---
title: Application.FileDialog property (Publisher)
keywords: vbapb10.chm131089
f1_keywords:
- vbapb10.chm131089
ms.prod: publisher
api_name:
- Publisher.Application.FileDialog
ms.assetid: 65d73a9d-be4c-d809-d10d-468181ef9eb0
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.FileDialog property (Publisher)

Returns a **[FileDialog](office.filedialog.md)** object that represents a single instance of a file dialog box.


## Syntax

_expression_.**FileDialog** (_Type_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Required| **[MsoFileDialogType](Office.MsoFileDialogType.md)**| The type of dialog box.|

## Return value

FileDialog


## Remarks

The _Type_ parameter can be one of the **MsoFileDialogType** constants declared in the Microsoft Office type library.


## Example

This example displays the **Save As** dialog box and stores the file name specified by the user.

```vb
Sub ShowSaveAsDialog() 
 Dim dlgSaveAs As FileDialog 
 Dim strFile As String 
 
 Set dlgSaveAs = Application.FileDialog( _ 
 Type:=msoFileDialogSaveAs) 
 dlgSaveAs.Show 
 strFile = dlgSaveAs.SelectedItems(1) 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]