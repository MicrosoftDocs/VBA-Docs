---
title: Worksheet.PivotTableUpdate Event (Excel)
keywords: vbaxl10.chm502081
f1_keywords:
- vbaxl10.chm502081
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableUpdate
ms.assetid: 66186c97-6855-b360-a6c0-56da617d24a6
ms.date: 06/08/2017
localization_priority: Priority
---


# Worksheet.PivotTableUpdate Event (Excel)

Occurs after a PivotTable report is updated on a worksheet.


## Syntax

_expression_. `PivotTableUpdate`( `_Target_` )

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|

## Return value

Nothing


## Example

This example displays a message stating that the PivotTable report has been updated. This example assumes you have declared an object of type  **Worksheet** with events in a class module.


```vb
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been updated." 
 
End Sub
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]