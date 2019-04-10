---
title: Has3DEffect property (Excel Graph)
keywords: vbagr10.chm67201
f1_keywords:
- vbagr10.chm67201
ms.prod: excel
api_name:
- Excel.Has3DEffect
ms.assetid: e19f4d47-ca7b-ea70-01eb-ced3c1dd343f
ms.date: 04/11/2019
localization_priority: Normal
---


# Has3DEffect property (Excel Graph)

**True** if the series has a three-dimensional appearance. Applies only to bubble charts. Read/write **Boolean**.

## Syntax

_expression_.**Has3DEffect**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example gives series one on the bubble chart a three-dimensional appearance.

```vb
With myChart 
 .SeriesCollection(1).Has3DEffect = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]