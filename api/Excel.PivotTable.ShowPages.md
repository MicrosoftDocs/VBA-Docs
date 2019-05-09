---
title: PivotTable.ShowPages method (Excel)
keywords: vbaxl10.chm235077
f1_keywords:
- vbaxl10.chm235077
ms.prod: excel
api_name:
- Excel.PivotTable.ShowPages
ms.assetid: 7ebb55ab-ecda-31f7-23d2-fdefc12ee161
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.ShowPages method (Excel)

Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.


## Syntax

_expression_.**ShowPages** (_PageField_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PageField_|Optional| **Variant**|A string that names a single page field in the report.|

## Return value

Variant


## Remarks

This method isn't available for OLAP data sources.


## Example

This example creates a new PivotTable report for each item in the page field (the field named Country).

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.ShowPages "Country"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]