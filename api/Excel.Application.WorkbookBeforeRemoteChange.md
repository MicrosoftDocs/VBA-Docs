---
title: Application.WorkbookBeforeRemoteChange event (Excel)
keywords: vbaxl10.chm503113
f1_keywords:
- vbaxl10.chm503113
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeRemoteChange
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookBeforeRemoteChange event (Excel)

Occurs before a remote user's edits to the workbook are merged.

## Syntax

_expression_.**WorkbookBeforeRemoteChange** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook that has been changed by a remote user.|

## Return value

Nothing

## Example

This example shows you where you can place code that runs before merging an incoming remote change. This code must be placed in a class module and an instance of that class must be correctly initialized. 

For more information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).

```vb
Private Sub App_WorkbookBeforeRemoteChange(ByVal Wb As Workbook)
    'A remote user has made a change to this workbook.
    'The code in this subroutine will be run before those changes are merged.
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]