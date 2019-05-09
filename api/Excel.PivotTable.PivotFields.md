---
title: PivotTable.PivotFields method (Excel)
keywords: vbaxl10.chm235089
f1_keywords:
- vbaxl10.chm235089
ms.prod: excel
api_name:
- Excel.PivotTable.PivotFields
ms.assetid: 2729eef0-bfe6-1683-8bb1-f12d8d03d939
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PivotFields method (Excel)

Returns an object that represents either a single PivotTable field (a **[PivotField](Excel.PivotField.md)** object) or a collection of both the visible and hidden fields (a **[PivotFields](Excel.PivotFields.md)** object) in the PivotTable report. Read-only.


## Syntax

_expression_.**PivotFields** (_Index_)

_expression_ An expression that returns a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned.|

## Return value

Object


## Remarks

For OLAP data sources, there are no hidden fields, and the object or collection that is returned reflects what's currently visible.


## Example

This example adds the PivotTable report's field names to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.PivotFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
