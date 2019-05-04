---
title: PivotCell.Range property (Excel)
keywords: vbaxl10.chm692080
f1_keywords:
- vbaxl10.chm692080
ms.prod: excel
api_name:
- Excel.PivotCell.Range
ms.assetid: b0b52ca0-a73b-acc3-25a8-330da27e4f92
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.Range property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range that the specified PivotCell applies to.


## Syntax

_expression_.**Range**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Example

The following example stores in a variable the address for the AutoFilter applied to the Crew worksheet.

```vb
rAddress = Worksheets("Crew").AutoFilter.Range.Address
```

<br/>

This example scrolls through the workbook window until the hyperlink range is in the upper-left corner of the active window.

```vb
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).Range 
ActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]