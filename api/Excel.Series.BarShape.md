---
title: Series.BarShape property (Excel)
keywords: vbaxl10.chm578114
f1_keywords:
- vbaxl10.chm578114
ms.prod: excel
api_name:
- Excel.Series.BarShape
ms.assetid: 27af7eea-6ad3-e906-c5f8-d9e29314b32d
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.BarShape property (Excel)

Returns or sets the shape used with the 3D bar or column chart. Read/write **[XlBarShape](Excel.XlBarShape.md)**.


## Syntax

_expression_.**BarShape**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example sets the shape used with series one on chart one.

```vb
Charts(1).SeriesCollection(1).BarShape = xlConeToPoint
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]