---
title: Range.AutoOutline method (Excel)
keywords: vbaxl10.chm144087
f1_keywords:
- vbaxl10.chm144087
ms.prod: excel
api_name:
- Excel.Range.AutoOutline
ms.assetid: a2553695-6d45-9b7c-7c45-5255fa3641f0
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.AutoOutline method (Excel)

Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.


## Syntax

_expression_.**AutoOutline**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Example

This example creates an outline for the range A1:G37 on Sheet1. 

> [!NOTE] 
> The range must contain either a **summary row** or a **summary column**.


```vb
Worksheets("Sheet1").Range("A1:G37").AutoOutline
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]