---
title: Application.WorkbookPivotTableCloseConnection event (Excel)
keywords: vbaxl10.chm504095
f1_keywords:
- vbaxl10.chm504095
ms.prod: excel
api_name:
- Excel.Application.WorkbookPivotTableCloseConnection
ms.assetid: 4c1d4cb2-f589-3c3c-ab4c-dcb08467fcfb
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookPivotTableCloseConnection event (Excel)

Occurs after a PivotTable report connection has been closed.


## Syntax

_expression_.**WorkbookPivotTableCloseConnection** (_Wb_, _Target_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The selected workbook.|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|

## Return value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been closed. This example assumes that you have declared an object of type **Workbook** with events in a class module.

```vb
Private Sub ConnectionApp_WorkbookPivotTableCloseConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]