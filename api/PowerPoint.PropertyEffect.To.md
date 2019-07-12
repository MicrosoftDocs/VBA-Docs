---
title: PropertyEffect.To property (PowerPoint)
keywords: vbapp10.chm662006
f1_keywords:
- vbapp10.chm662006
ms.prod: powerpoint
api_name:
- PowerPoint.PropertyEffect.To
ms.assetid: 453cc64b-88b7-e543-fff5-d218b8cc320f
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyEffect.To property (PowerPoint)

Sets or returns a  **Variant** that represents the ending value of an object's property. Read/write.


## Syntax

_expression_. `To`

_expression_ A variable that represents a [PropertyEffect](PowerPoint.PropertyEffect.md) object.


## Return value

Variant


## Remarks

The default value is  **Empty**, in which case the current position of the object is used.

Do not confuse this property with the  **ToX** or **ToY** properties of the **[ScaleEffect](PowerPoint.ScaleEffect.md)** and **[MotionEffect](PowerPoint.MotionEffect.md)** objects, which are only used for scaling or motion effects.


## See also


[PropertyEffect Object](PowerPoint.PropertyEffect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]