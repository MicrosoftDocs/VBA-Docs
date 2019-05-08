---
title: Point.ApplyPictToFront property (Excel)
keywords: vbaxl10.chm576095
f1_keywords:
- vbaxl10.chm576095
ms.prod: excel
api_name:
- Excel.Point.ApplyPictToFront
ms.assetid: e739e368-9789-be23-da90-17ab4cf3a935
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.ApplyPictToFront property (Excel)

**True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean**.


## Syntax

_expression_.**ApplyPictToFront**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example applies pictures to the front of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).

```vb
Charts(1).SeriesCollection(1).ApplyPictToFront = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]