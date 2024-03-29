---
title: ApplyPictToFront property (Excel Graph)
keywords: vbagr10.chm3076954
f1_keywords:
- vbagr10.chm3076954
api_name:
- Excel.ApplyPictToFront
ms.assetid: c6b1b61c-edb1-fb40-387a-0106e8ca225a
ms.date: 04/09/2019
ms.localizationpriority: medium
---


# ApplyPictToFront property (Excel Graph)

**True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean**.

## Syntax

_expression_.**ApplyPictToFront**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example applies pictures to the front of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).

```vb
myChart.SeriesCollection(1).ApplyPictToFront = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]