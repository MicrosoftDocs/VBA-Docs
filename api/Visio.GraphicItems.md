---
title: GraphicItems object (Visio)
keywords: vis_sdr.chm61030
f1_keywords:
- vis_sdr.chm61030
ms.prod: visio
api_name:
- Visio.GraphicItems
ms.assetid: 89d0bbeb-ee45-50cc-490e-0af49d036ad1
ms.date: 06/19/2019
localization_priority: Normal
---


# GraphicItems object (Visio)

The collection of **[GraphicItem](Visio.GraphicItem.md)** objects associated with a **[Master](Visio.Master.md)** object of type **visTypeDataGraphic**.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the **GraphicItems** collection is **Item**.

The **AddCopy** method adds a copy of an existing **GraphicItem** object to the **GraphicItems** collection. The **GraphicItem** object to be added must already exist in another master of type **visTypeDataGraphic**.

Use the **DataGraphic** property to return the **Master** object of type **visTypeDataGraphic** that the **GraphicItems** collection is associated with.

> [!NOTE] 
> You must create masters of type **visTypeDataGraphic** by using the Visio user interface—you cannot create them programmatically. For more information about creating these masters, search for "data graphics" in Visio end-user Help.

## Methods

-  [AddCopy](Visio.GraphicItems.AddCopy.md)

## Properties

-  [Application](Visio.GraphicItems.Application.md)
-  [Count](Visio.GraphicItems.Count.md)
-  [DataGraphic](Visio.GraphicItems.DataGraphic.md)
-  [Document](Visio.GraphicItems.Document.md)
-  [Item](Visio.GraphicItems.Item.md)
-  [ItemFromID](Visio.GraphicItems.ItemFromID.md)
-  [ObjectType](Visio.GraphicItems.ObjectType.md)
-  [Stat](Visio.GraphicItems.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]