---
title: FileDialogFilters.Clear method (Office)
keywords: vbaof11.chm255005
f1_keywords:
- vbaof11.chm255005
ms.prod: office
api_name:
- Office.FileDialogFilters.Clear
ms.assetid: 1d5fa55e-6a61-d808-51a4-86116420f89f
ms.date: 01/09/2019
localization_priority: Normal
---


# FileDialogFilters.Clear method (Office)

Removes all filters currently applied in a file dialog box.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a **[FileDialogFilters](Office.FileDialogFilters.md)** object.


## Example

The following example iterates through the default filters of the **SaveAs** dialog box and displays the description of each filter that includes a Microsoft Excel file.


```vb
Sub Main() 
 
 'Declare a variable as a FileDialogFilters collection. 
 Dim fdfs As FileDialogFilters 
 
 'Declare a variable as a FileDialogFilter object. 
 Dim fdf As FileDialogFilter 
 
 'Set the FileDialogFilters collection variable to 
 'the FileDialogFilters collection of the SaveAs dialog box. 
 Set fdfs = Application.FileDialog(msoFileDialogSaveAs).Filters 
 
 'Iterate through the description and extensions of each 
 'default filter in the SaveAs dialog box. 
 For Each fdf In fdfs 
 
 'Display the description of filters that include 
 'Microsoft Excel files 
 If InStr(1, fdf.Extensions, "xls", vbTextCompare) > 0 Then 
 MsgBox "Description of filter: " & fdf.Description 
 End If 
 Next fdf 
 
End Sub
```


## See also

- [FileDialogFilters object members](overview/library-reference/filedialogfilters-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]