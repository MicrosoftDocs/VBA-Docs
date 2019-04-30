---
title: PivotTable.Tag property (Excel)
keywords: vbaxl10.chm235127
f1_keywords:
- vbaxl10.chm235127
ms.prod: excel
api_name:
- Excel.PivotTable.Tag
ms.assetid: 7ef25e2e-6c89-3654-4045-2937fcf47121
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotTable.Tag property (Excel)

Returns or sets a string saved with the PivotTable report. Read/write  **String**.


## Syntax

_expression_.**Tag**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example sets the PivotTable report's  **Tag** property.


```vb
Worksheets(1).PivotTables("Pivot1").Tag = "Product Sales by Region"
```


## See also


[PivotTable Object](Excel.PivotTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]