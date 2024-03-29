---
title: ChartObject.Duplicate method (Excel)
keywords: vbaxl10.chm494080
f1_keywords:
- vbaxl10.chm494080
api_name:
- Excel.ChartObject.Duplicate
ms.assetid: f43de123-c113-ca5d-6cf7-71f7d08f7e88
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# ChartObject.Duplicate method (Excel)

Duplicates the object and returns a reference to the new copy.


## Syntax

_expression_.**Duplicate**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Return value

Object


## Example

This example duplicates embedded chart one on Sheet1 and then selects the copy.

```vb
Set dChart = Worksheets("Sheet1").ChartObjects(1).Duplicate 
dChart.Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]