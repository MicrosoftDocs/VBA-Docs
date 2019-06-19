---
title: Path object (Visio)
keywords: vis_sdr.chm10200
f1_keywords:
- vis_sdr.chm10200
ms.prod: visio
api_name:
- Visio.Path
ms.assetid: 6bdbbd2f-e375-bb9d-87e3-c4d8997d2aab
ms.date: 06/19/2019
localization_priority: Normal
---


# Path object (Visio)

Represents a sequence of one or more segments whose ends abut. A path describes where a pen would move in order to draw one shape component. Each **Path** object corresponds to a Geometry section of a shape.


## Remarks

The default property of a **Path** object is **Item**.

A **[Curve](visio.curve.md)** object is an item in a **Path** object that is any linear or curved segment representing a consecutive sequence of rows in the Geometry section that the **Path** object represents. The number of **Curve** objects in a **Path** object is not necessarily the same as the number of rows in its Geometry section.

The **Path** object is conceptually of zero width. Line weights, patterns, and ends are ignored, but corner rounding is included. A **Path** object may or may not be closed, and it may intersect itself. For example, a **Path** object may describe a figure eight.

If you retrieve a **Path** object from a collection obtained by the **Paths** property of a shape, its coordinates are expressed in the shape's parent coordinate system. 

If you retrieve a **Path** object from a collection obtained by the **PathsLocal** property of a shape, its coordinates are expressed in the shape's local coordinate system. In both cases, coordinates are expressed in internal drawing units (inches).

## Methods

-  [Points](Visio.Path.Points.md)

## Properties

-  [Application](Visio.Path.Application.md)
-  [Closed](Visio.Path.Closed.md)
-  [Count](Visio.Path.Count.md)
-  [Item](Visio.Path.Item.md)
-  [ObjectType](Visio.Path.ObjectType.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]