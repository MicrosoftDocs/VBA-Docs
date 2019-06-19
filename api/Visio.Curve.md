---
title: Curve object (Visio)
keywords: vis_sdr.chm10075
f1_keywords:
- vis_sdr.chm10075
ms.prod: visio
api_name:
- Visio.Curve
ms.assetid: 040f47b2-794d-72c7-7479-b61d8f1cb75f
ms.date: 06/19/2019
localization_priority: Normal
---


# Curve object (Visio)

An item in a **[Path](visio.path.md)** object that represents a consecutive sequence of rows in the Geometry section of its **Path** object.


## Remarks

The default property of a **Curve** object is **Point**.

If a **Curve** object is in a collection returned by the **[Paths](visio.shape.paths.md)** property of a **Shape** object, its coordinates are expressed in the shape's parent coordinate system. 

If the **Curve** object is in a collection returned by the **[PathsLocal](visio.shape.pathslocal.md)** property of a **Shape** object, its coordinates are expressed in the shape's local coordinate system. In both cases, the coordinates are expressed in internal drawing units (inches).

A **Curve** object describes itself in terms of its parameter domain, which is the range `[Start(),End()]`. 

Use the **Start** property to obtain the curve's starting point and the **End** property to obtain the curve's ending point.

Use the **Point** method to extrapolate a point along the curve's path. 

Use the **PointAndDerivatives** method to determine a point along the curve's path and, optionally, its first and second derivatives.

Use the **Points** method to obtain a stream of points that approximate the curve's path.

## Methods

-  [Point](Visio.Curve.Point.md)
-  [PointAndDerivatives](Visio.Curve.PointAndDerivatives.md)
-  [Points](Visio.Curve.Points.md)

## Properties

-  [Application](Visio.Curve.Application.md)
-  [Closed](Visio.Curve.Closed.md)
-  [End](Visio.Curve.End.md)
-  [ObjectType](Visio.Curve.ObjectType.md)
-  [Start](Visio.Curve.Start.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]