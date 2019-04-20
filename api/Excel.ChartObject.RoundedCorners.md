---
title: ChartObject.RoundedCorners property (Excel)
keywords: vbaxl10.chm494101
f1_keywords:
- vbaxl10.chm494101
ms.prod: excel
api_name:
- Excel.ChartObject.RoundedCorners
ms.assetid: cb58389a-0235-384e-e32a-e669e789bacc
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.RoundedCorners property (Excel)

**True** if the embedded chart has rounded corners. Read/write **Boolean**.


## Syntax

_expression_.**RoundedCorners**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example adds rounded corners to embedded chart one on Sheet1.

```vb
Worksheets("Sheet1").ChartObjects(1).RoundedCorners = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]