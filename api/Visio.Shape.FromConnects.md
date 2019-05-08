---
title: Shape.FromConnects property (Visio)
keywords: vis_sdr.chm11213580
f1_keywords:
- vis_sdr.chm11213580
ms.prod: visio
api_name:
- Visio.Shape.FromConnects
ms.assetid: feb80221-c5d9-f72e-2f79-5153ed375383
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.FromConnects property (Visio)

Returns a  **Connects** collection of the shapes connected to a shape. Read-only.


## Syntax

_expression_. `FromConnects`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

Connects


## Remarks

The  **FromConnects** property of a shape returns a **Connects** collection that contains every **Connect** object for which the shape is the **ToSheet** property. This tells you all the shapes connected to a shape.

To obtain a  **Connects** collection that contains every **Connect** object for which the shape is the **FromSheet** property, use the shape's **Connects** property. This tells you all the shapes to which the shape is connected.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]