---
title: PropertyEffect.From property (PowerPoint)
keywords: vbapp10.chm662005
f1_keywords:
- vbapp10.chm662005
ms.prod: powerpoint
api_name:
- PowerPoint.PropertyEffect.From
ms.assetid: 314435d3-27aa-323b-65f4-de7f7864f30d
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyEffect.From property (PowerPoint)

Sets or returns a **Variant** that represents the starting value of an object's property. Read/write.


## Syntax

_expression_. `From`

_expression_ A variable that represents a [PropertyEffect](PowerPoint.PropertyEffect.md) object.


## Remarks

the **From** property is similar to the **[Points](PowerPoint.PropertyEffect.Points.md)** property, but using the **From** property is easier for simple tasks.

The default value is **Empty**, in which case the current position of the object is used.

Do not confuse this property with the **FromX** or **FromY** properties of the **[ScaleEffect](PowerPoint.ScaleEffect.md)** and **[MotionEffect](PowerPoint.MotionEffect.md)** objects, which are only used for scaling or motion effects.


## See also


[PropertyEffect Object](PowerPoint.PropertyEffect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]