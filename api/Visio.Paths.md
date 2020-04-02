---
title: Paths object (Visio)
keywords: vis_sdr.chm10205
f1_keywords:
- vis_sdr.chm10205
ms.prod: visio
api_name:
- Visio.Paths
ms.assetid: 9adcc130-555e-7eee-d9a0-66ee7116e41f
ms.date: 06/19/2019
localization_priority: Normal
---


# Paths object (Visio)

Includes a **[Path](Visio.Path.md)** object for each Geometry section for a group or shape.


## Remarks

To retrieve a **Paths** collection expressed in the shape's parent coordinate system, use the **[Paths](visio.shape.paths.md)** property of the shape. The coordinates are expressed in internal drawing units (inches).

The default property of a **Paths** collection is **Item**.

To retrieve a **Paths** collection expressed in the shape's local coordinate system, use the **[PathsLocal](visio.shape.pathslocal.md)** property of the shape. The coordinates are expressed in internal drawing units (inches).

If a **Shape** object is a page, foreign object, or guide, its **Paths** and **PathsLocal** properties don't contain any items.

If a **Shape** object is a group, its **Paths** and **PathsLocal** properties are the union of the paths of its component shapes.

If a **Shape** object is a shape, its **Paths** and **PathsLocal** properties include one item for each Geometry section that defines a stroke of positive length.

## Properties

- [Application](Visio.Paths.Application.md)
- [Count](Visio.Paths.Count.md)
- [Item](Visio.Paths.Item.md)
- [ObjectType](Visio.Paths.ObjectType.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]