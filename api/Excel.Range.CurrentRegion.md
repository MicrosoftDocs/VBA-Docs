---
title: Range.CurrentRegion property (Excel)
keywords: vbaxl10.chm144111
f1_keywords:
- vbaxl10.chm144111
ms.prod: excel
api_name:
- Excel.Range.CurrentRegion
ms.assetid: 39277cc5-07ff-8453-7330-b272b365f9dc
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.CurrentRegion property (Excel)

Returns a **Range** object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. Read-only.


## Syntax

_expression_.**CurrentRegion**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

This property is useful for many operations that automatically expand the selection to include the entire current region, such as the **[XlRangeAutoFormat](Excel.XlRangeAutoFormat.md)** values.

This property cannot be used on a protected worksheet.


## Example

This example selects the current region on Sheet1.

```vb
Worksheets("Sheet1").Activate 
ActiveCell.CurrentRegion.Select
```

<br/>

This example assumes that you have a table on Sheet1 that has a header row. The example selects the table without selecting the header row. The active cell must be somewhere in the table before you run the example.

```vb
Set tbl = ActiveCell.CurrentRegion 
tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, _ 
 tbl.Columns.Count).Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
