---
title: Worksheet.PivotTables method (Excel)
keywords: vbaxl10.chm175118
f1_keywords:
- vbaxl10.chm175118
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTables
ms.assetid: b60944cd-827d-15dc-d49e-c739c237de15
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PivotTables method (Excel)

Returns an object that represents either a single PivotTable report (a **[PivotTable](Excel.PivotTable.md)** object) or a collection of all the PivotTable reports (a **[PivotTables](Excel.PivotTables.md)** object) on a worksheet. Read-only.


## Syntax

_expression_.**PivotTables** (_Index_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the report.|

## Return value

**Object**


## Example

This example sets the Sum of 1994 field in the first PivotTable report on the active sheet to use the SUM function.

```vb
ActiveSheet.PivotTables("PivotTable1"). _ 
 PivotFields("Sum of 1994").Function = xlSum
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
