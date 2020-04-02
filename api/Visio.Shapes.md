---
title: Shapes object (Visio)
keywords: vis_sdr.chm10230
f1_keywords:
- vis_sdr.chm10230
ms.prod: visio
api_name:
- Visio.Shapes
ms.assetid: 9ec3c379-54c2-50d8-4f6b-79a95b8d12f0
ms.date: 06/19/2019
localization_priority: Normal
---


# Shapes object (Visio)

Includes a **[Shape](Visio.Shape.md)** object for each basic shape, group, guide, or object from another application (linked or embedded in Microsoft Visio) on a drawing page, master, or group.


## Remarks

To retrieve a **Shapes** collection, use the **Shapes** property of a **Page**, **Master**, or **Shape** object.

The default property of a **Shapes** collection is **Item**.

The order of items in a **Shapes** collection corresponds to the stacking (drawing) order of the shapes.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this collection maps to the following types:

- **Microsoft.Office.Interop.Visio.IVShapes**
    

## Methods

- [CenterDrawing](Visio.Shapes.CenterDrawing.md)

## Properties

- [Application](Visio.Shapes.Application.md)
- [ContainingMaster](Visio.Shapes.ContainingMaster.md)
- [ContainingPage](Visio.Shapes.ContainingPage.md)
- [ContainingShape](Visio.Shapes.ContainingShape.md)
- [Count](Visio.Shapes.Count.md)
- [Document](Visio.Shapes.Document.md)
- [EventList](Visio.Shapes.EventList.md)
- [Item](Visio.Shapes.Item.md)
- [ItemFromID](Visio.Shapes.ItemFromID.md)
- [ItemFromUniqueID](Visio.Shapes.ItemFromUniqueID.md)
- [ItemU](Visio.Shapes.ItemU.md)
- [ObjectType](Visio.Shapes.ObjectType.md)
- [PersistsEvents](Visio.Shapes.PersistsEvents.md)
- [Stat](Visio.Shapes.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]