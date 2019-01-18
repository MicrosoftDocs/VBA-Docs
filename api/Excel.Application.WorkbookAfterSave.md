---
title: Application.WorkbookAfterSave Event (Excel)
keywords: vbaxl10.chm504114
f1_keywords:
- vbaxl10.chm504114
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterSave
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookAfterSave Event (Excel)

Occurs after the workbook is saved.

**NOTE:** In Office 365, Excel supports AutoSave, which enables the user's edits to be saved automatically and continuously. Following the guidance in [this article](../Library-Reference/Concepts/how-autosave-impacts-addins-and-macros.md) will ensure that running code in response to the **WorkbookAfterSave** event will function as intended when AutoSave is enabled.

## Syntax

_expression_. `WorkbookAfterSave`( `_Wb_` , `_Success_` )

_expression_ A variable that represents an '[Application](Excel.Application(object).md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook being saved.|
| _Success_|Required| **Boolean**|Returns  **True** if the save operation was successful; otherwise **False**.|

## Return value

Nothing


## Example

The following code example displays a message box if the workbook was successfully saved. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the  **Application** object, see [Using Events with the Application Object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " & Wb.Name & " workbook was successfully saved.") 
End If 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

[AutoSave](../Library-Reference/Concepts/how-autosave-impacts-addins-and-macros.md)
