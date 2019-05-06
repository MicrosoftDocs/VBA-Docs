---
title: PivotField.VisibleItems property (Excel)
keywords: vbaxl10.chm240099
f1_keywords:
- vbaxl10.chm240099
ms.prod: excel
api_name:
- Excel.PivotField.VisibleItems
ms.assetid: f5c0f367-42a4-fffe-5b27-af2c19890ad3
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.VisibleItems property (Excel)

Returns an object that represents either a single visible PivotTable item (a **[PivotItem](Excel.PivotItem.md)** object) or a collection of all the visible items (a **[PivotItems](Excel.PivotItems.md)** object) in the specified field. Read-only.


## Syntax

_expression_.**VisibleItems** (_Index_)

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number or name of the item to be returned (can be an array to specify more than one item).|

## Remarks

For OLAP data sources, this property is read-only and always returns **True**. There are no hidden items.


## Example

This example adds the names of all visible items in the field named Product to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("Product").VisibleItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]