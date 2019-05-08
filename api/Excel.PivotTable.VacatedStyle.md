---
title: PivotTable.VacatedStyle property (Excel)
keywords: vbaxl10.chm235129
f1_keywords:
- vbaxl10.chm235129
ms.prod: excel
api_name:
- Excel.PivotTable.VacatedStyle
ms.assetid: 94be037f-3fce-ad39-9dd6-b72f829c3fbf
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.VacatedStyle property (Excel)

Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default). Read/write **String**.


## Syntax

_expression_.**VacatedStyle**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example sets the vacated cells in the PivotTable report to the BlackAndBlue style.

```vb
Worksheets(1).PivotTables("Pivot1").VacatedStyle = "BlackAndBlue"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]