---
title: MenuSets object (Visio)
keywords: vis_sdr.chm10175
f1_keywords:
- vis_sdr.chm10175
ms.prod: visio
api_name:
- Visio.MenuSets
ms.assetid: 6a49d679-abdb-2bd4-134b-c61ea3f196e8
ms.date: 06/19/2019
localization_priority: Normal
---


# MenuSets object (Visio)

Includes a **[MenuSet](Visio.MenuSet.md)** object for each Microsoft Visio window context that has menus.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **MenuSets** collection, use the **[MenuSets](visio.uiobject.menusets.md)** property of a **UIObject** object.

The default property of a **MenuSets** collection is **Item**.

Unlike other Visio collections, the **MenuSets** collection is indexed starting with zero (0) rather than 1.

A **MenuSet** object is identified in the **MenuSets** collection by its **[SetID](Visio.MenuSet.SetID.md)** property, which corresponds to a Visio window context. For a list of **SetID** values for **MenuSet** objects, see the **SetID** property.

## Methods

-  [Add](Visio.MenuSets.Add.md)
-  [AddAtID](Visio.MenuSets.AddAtID.md)

## Properties

-  [Count](Visio.MenuSets.Count.md)
-  [Item](Visio.MenuSets.Item.md)
-  [ItemAtID](Visio.MenuSets.ItemAtID.md)
-  [Parent](Visio.MenuSets.Parent.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]