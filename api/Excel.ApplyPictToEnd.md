---
title: ApplyPictToEnd property (Excel Graph)
keywords: vbagr10.chm67197
f1_keywords:
- vbagr10.chm67197
ms.prod: excel
api_name:
- Excel.ApplyPictToEnd
ms.assetid: a063278c-9dc5-a28e-49c7-3045b8927c2e
ms.date: 04/09/2019
localization_priority: Normal
---


# ApplyPictToEnd property (Excel Graph)

**True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean**.

## Syntax

_expression_.**ApplyPictToEnd**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example applies pictures to the end of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).

```vb
myChart.SeriesCollection(1).ApplyPictToEnd = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]