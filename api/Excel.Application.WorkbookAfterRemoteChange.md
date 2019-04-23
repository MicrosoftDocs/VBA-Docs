---
title: Application.WorkbookAfterRemoteChange event (Excel)
keywords: vbaxl10.chm503114
f1_keywords:
- vbaxl10.chm503114
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterRemoteChange
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookAfterRemoteChange event (Excel)

Occurs after a remote user's edits to the workbook are merged.

## Syntax

_expression_.**WorkbookAfterRemoteChange** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook which has been changed by a remote user.|

## Return value

Nothing

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

This example shows you where you can place code that runs after merging an incoming remote change. This code must be placed in a class module, and an instance of that class must be correctly initialized.

```vb
Private Sub App_WorkbookAfterRemoteChange(ByVal Wb As Workbook)
    'A remote user has made a change to this workbook and that change has been merged.
    'The code in this subroutine will now be run.
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]