---
title: Interior.Pattern property (Excel)
keywords: vbaxl10.chm551076
f1_keywords:
- vbaxl10.chm551076
ms.prod: excel
api_name:
- Excel.Interior.Pattern
ms.assetid: 90587a6d-273c-00df-bb12-1a4415591705
ms.date: 06/08/2017
localization_priority: Normal
---


# Interior.Pattern property (Excel)

Returns or sets a  **Variant** value, containing an **[XlPattern](Excel.XlPattern.md)** constant, that represents the interior pattern.


## Syntax

_expression_.**Pattern**

_expression_ A variable that represents an [Interior](Excel.Interior-graph-property.md) object.


## Example

This example adds a crisscross pattern to the interior of cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1"). _ 
 Interior.Pattern = xlPatternCrissCross
```


## See also


[Interior Object](Excel.Interior(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
