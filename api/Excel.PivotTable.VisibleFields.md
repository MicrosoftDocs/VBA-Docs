---
title: PivotTable.VisibleFields property (Excel)
keywords: vbaxl10.chm235101
f1_keywords:
- vbaxl10.chm235101
ms.prod: excel
api_name:
- Excel.PivotTable.VisibleFields
ms.assetid: 01d5e76d-e109-905d-1743-1fbacd85e7a6
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.VisibleFields property (Excel)

Returns an object that represents either a single field in a PivotTable report (a **[PivotField](Excel.PivotField.md)** object) or a collection of all the visible fields (a **[PivotFields](Excel.PivotFields.md)** object). Visible fields are shown as row, column, page or data fields. Read-only.

## Syntax

_expression_.**VisibleFields** (_Index_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned (can be an array to specify more than one field).|

## Remarks

For OLAP data sources, there are no hidden fields, and this property returns all the fields in the PivotTable cache.


## Example

This example adds the visible field names to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.VisibleFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]