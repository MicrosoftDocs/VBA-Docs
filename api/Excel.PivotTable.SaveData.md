---
title: PivotTable.SaveData property (Excel)
keywords: vbaxl10.chm235096
f1_keywords:
- vbaxl10.chm235096
ms.prod: excel
api_name:
- Excel.PivotTable.SaveData
ms.assetid: f8f788cf-b8a2-4694-1a52-f48e00e6471c
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.SaveData property (Excel)

**True** if data for the PivotTable report is saved with the workbook. **False** if only the report definition is saved. Read/write **Boolean**.


## Syntax

_expression_.**SaveData**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For OLAP data sources, this property is always set to **False**.


## Example

This example sets the PivotTable report to save data with the workbook.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.SaveData = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]