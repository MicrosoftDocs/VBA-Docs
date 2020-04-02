---
title: ToolbarSets object (Visio)
keywords: vis_sdr.chm10295
f1_keywords:
- vis_sdr.chm10295
ms.prod: visio
api_name:
- Visio.ToolbarSets
ms.assetid: ddf79048-6585-81ab-b1c6-d7c4b0f0ff1b
ms.date: 06/19/2019
localization_priority: Normal
---


# ToolbarSets object (Visio)

Includes a **[ToolbarSet](Visio.ToolbarSet.md)** object for each window context that can display toolbars.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **ToolbarSets** collection, use the **[ToolbarSets](visio.uiobject.toolbarsets.md)** property of a **UIObject** object.

The default property of a **ToolbarSets** collection is **Item**.

Unlike other Microsoft Visio collections, the **ToolbarSets** collection is indexed starting with zero (0) rather than 1.

A **ToolbarSet** object is identified in the **ToolbarSets** collection by its **[SetID](Visio.ToolbarSet.SetID.md)** property, which corresponds to a Visio window context. For a list of **SetID** values for **ToolbarSet** objects, see the **SetID** property.

## Methods

- [Add](Visio.ToolbarSets.Add.md)
- [AddAtID](Visio.ToolbarSets.AddAtID.md)

## Properties

- [Count](Visio.ToolbarSets.Count.md)
- [Item](Visio.ToolbarSets.Item.md)
- [ItemAtID](Visio.ToolbarSets.ItemAtID.md)
- [Parent](Visio.ToolbarSets.Parent.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]