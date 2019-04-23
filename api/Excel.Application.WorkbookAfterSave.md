---
title: Application.WorkbookAfterSave event (Excel)
keywords: vbaxl10.chm504114
f1_keywords:
- vbaxl10.chm504114
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterSave
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookAfterSave event (Excel)

Occurs after the workbook is saved.

> [!NOTE] 
> In Office 365, Excel supports AutoSave, which enables the user's edits to be saved automatically and continuously. For more information, see [How AutoSave impacts add-ins and macros](../Library-Reference/Concepts/how-autosave-impacts-addins-and-macros.md) to ensure that running code in response to the **WorkbookAfterSave** event functions as intended when AutoSave is enabled.

## Syntax

_expression_.**WorkbookAfterSave** (_Wb_, _Success_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook being saved.|
| _Success_|Required| **Boolean**|Returns **True** if the save operation was successful; otherwise, **False**.|

## Return value

Nothing

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

The following code example displays a message box if the workbook was successfully saved. This code must be placed in a class module, and an instance of that class must be correctly initialized. 

```vb
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " & Wb.Name & " workbook was successfully saved.") 
End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]