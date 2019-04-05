---
title: Application.WorkbookPivotTableOpenConnection event (Excel)
keywords: vbaxl10.chm504096
f1_keywords:
- vbaxl10.chm504096
ms.prod: excel
api_name:
- Excel.Application.WorkbookPivotTableOpenConnection
ms.assetid: 5f07e995-96fd-86ac-2d1c-1366528fd8c6
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookPivotTableOpenConnection event (Excel)

Occurs after a PivotTable report connection has been opened.


## Syntax

_expression_.**WorkbookPivotTableOpenConnection** (_Wb_, _Target_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The selected workbook.|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|

## Return value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been opened. This example assumes that you have declared an object of type **Workbook** with events in a class module.

```vb
Private Sub ConnectionApp_WorkbookPivotTableOpenConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been opened." 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]