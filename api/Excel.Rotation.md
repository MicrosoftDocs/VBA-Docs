---
title: Rotation property (Excel Graph)
keywords: vbagr10.chm65595
f1_keywords:
- vbagr10.chm65595
ms.prod: excel
api_name:
- Excel.Rotation
ms.assetid: f78b6998-fae2-c80b-3a98-96ad359e6c47
ms.date: 04/12/2019
localization_priority: Normal
---


# Rotation property (Excel Graph)

Returns or sets the rotation of the 3D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3D charts. Read/write **Variant**.

## Syntax

_expression_.**Rotation**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the rotation of _myChart_ to 30 degrees. The example should be run on a 3D chart.

```vb
myChart.Rotation = 30
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]