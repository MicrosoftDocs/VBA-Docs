---
title: PivotField.HiddenItems property (Excel)
keywords: vbaxl10.chm240083
f1_keywords:
- vbaxl10.chm240083
ms.prod: excel
api_name:
- Excel.PivotField.HiddenItems
ms.assetid: ec30c18e-c030-23b8-2ea8-7ed7bfbd3312
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.HiddenItems property (Excel)

Returns an object that represents either a single hidden PivotTable item (a **[PivotItem](Excel.PivotItem.md)** object) or a collection of all the hidden items (a **[PivotItems](Excel.PivotItems.md)** object) in the specified field. Read-only.


## Syntax

_expression_.**HiddenItems** (_Index_)

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number or name of the item to be returned (can be an array to specify more than one item).|

## Remarks

For OLAP data sources, this property always returns an empty collection.


## Example

This example adds the names of all the hidden items in the field named Product to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("Product").HiddenItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]