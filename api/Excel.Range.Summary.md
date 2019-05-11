---
title: Range.Summary property (Excel)
keywords: vbaxl10.chm144207
f1_keywords:
- vbaxl10.chm144207
ms.prod: excel
api_name:
- Excel.Range.Summary
ms.assetid: f9e18651-20b6-1094-2ee5-7cd23559498e
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Summary property (Excel)

**True** if the range is an outlining summary row or column. The range should be a row or a column. Read-only **Variant**.


## Syntax

_expression_.**Summary**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example formats row four on Sheet1 as bold and italic if it's an outlining summary column.

```vb
With Worksheets("Sheet1").Rows(4) 
 If .Summary = True Then 
 .Font.Bold = True 
 .Font.Italic = True 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]