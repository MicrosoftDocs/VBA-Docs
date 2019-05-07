---
title: PivotField.PivotItems method (Excel)
keywords: vbaxl10.chm240091
f1_keywords:
- vbaxl10.chm240091
ms.prod: excel
api_name:
- Excel.PivotField.PivotItems
ms.assetid: 5ec5fa1e-a080-2cbf-e4d4-b15d39e13ac5
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.PivotItems method (Excel)

Returns an object that represents either a single PivotTable item (a **[PivotItem](Excel.PivotItem.md)** object) or a collection of all the visible and hidden items (a **[PivotItems](Excel.PivotItems.md)** object) in the specified field. Read-only.


## Syntax

_expression_.**PivotItems** (_Index_)

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the item to be returned.|

## Return value

Variant


## Remarks

For OLAP data sources, the collection is indexed by the unique name (the name returned by the **[SourceName](Excel.PivotField.SourceName.md)** property), not by the display name.


## Example

This example adds the names of all items in the field named Product to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtitem In pvtTable.PivotFields("Product").PivotItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtitem.Name 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
