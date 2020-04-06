---
title: Shape.Connects property (Visio)
keywords: vis_sdr.chm11213290
f1_keywords:
- vis_sdr.chm11213290
ms.prod: visio
api_name:
- Visio.Shape.Connects
ms.assetid: 9edaac59-f52e-67ee-6e5a-e11572907785
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Connects property (Visio)

Returns a  **Connects** collection for a shape, page, or master. Read-only.


## Syntax

_expression_. `Connects`

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

Connects


## Remarks

The  **Connects** collection of a shape contains every **Connect** object for which the shape is returned by the **FromSheet** property. This tells you all the shapes to which the shape is connected.

To obtain a  **Connects** collection that contains every **Connect** object for which the shape is the **ToSheet** property, use the shape's **FromConnects** property. This tells you all the shapes that are connected to this shape.

The  **Connects** collection of a page contains a **Connect** object for every connection on the page.

The  **Connects** collection of a master contains a **Connect** object for every connection in the master.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]