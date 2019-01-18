---
title: FileDialogFilter object (Office)
keywords: vbaof11.chm254000
f1_keywords:
- vbaof11.chm254000
ms.prod: office
api_name:
- Office.FileDialogFilter
ms.assetid: ff53a25a-0341-e761-01ef-6812ac9d64de
ms.date: 01/09/2019
localization_priority: Normal
---


# FileDialogFilter object (Office)

Represents a file filter in a file dialog box displayed through the **[FileDialog](office.filedialog.md)** object. Each file filter determines which files are displayed in the file dialog box.


## Remarks

Use the **[Item](office.filedialogfilters.item.md)** method with the **FileDialogFilters** collection to return a **FileDialogFilter** object. 

Use the **[Add](office.filedialogfilters.add.md)** method to add a **FileDialogFilter** object to the **FileDialogFilters** collection. 

You can return the extensions that a **FileDialogFilter** object uses to filter files with the **Extensions** property, and you can return the description of the filter with the **Description** property; however, both of these properties are read-only. If you want to set the extension or description, you must use the **Add** method.


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
 'Microsoft Excel files. 
 If InStr(1, fdf.Extensions, "xls", vbTextCompare) > 0 Then 
 MsgBox "Description of filter: " &amp; fdf.Description 
 End If 
 Next fdf 
End Sub
```


# See also

- [FileDialogFilter object members](overview/library-reference/filedialogfilter-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]