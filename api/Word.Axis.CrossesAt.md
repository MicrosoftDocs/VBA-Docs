---
title: Axis.CrossesAt property (Word)
keywords: vbawd10.chm113049608
f1_keywords:
- vbawd10.chm113049608
ms.prod: word
api_name:
- Word.Axis.CrossesAt
ms.assetid: 720fd3a6-89fb-bb55-9b0b-d6ecb2e5ca21
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.CrossesAt property (Word)

Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double**.


## Syntax

_expression_.**CrossesAt**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Setting this property causes the  **[Crosses](Word.Axis.Crosses.md)** property to change to **xlAxisCrossesCustom**. **xlAxisCrossesCustom** is a constant in the **xlAxisCrosses** enumeration.

You cannot use this property on radar charts. For 3D charts, this property indicates where the plane defined by the category axes crosses the value axis.


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]