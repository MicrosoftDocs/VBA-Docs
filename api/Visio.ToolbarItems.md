---
title: ToolbarItems object (Visio)
keywords: vis_sdr.chm10280
f1_keywords:
- vis_sdr.chm10280
ms.prod: visio
api_name:
- Visio.ToolbarItems
ms.assetid: 173cc711-7212-d56a-76a9-e30c3a608579
ms.date: 06/19/2019
localization_priority: Normal
---


# ToolbarItems object (Visio)

Includes a **[ToolbarItem](Visio.ToolbarItem.md)** object for each item on a toolbar.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **ToolbarItems** collection, use the **[ToolbarItems](Visio.Toolbar.ToolbarItems.md)** property of a **Toolbar** object.

The default property of a **ToolbarItems** collection is **Item**.

Unlike other Microsoft Visio collections, the **ToolbarItems** collection is indexed starting with zero (0) rather than 1.

## Methods

- [Add](Visio.ToolbarItems.Add.md)
- [AddAt](Visio.ToolbarItems.AddAt.md)

## Properties

- [Count](Visio.ToolbarItems.Count.md)
- [Item](Visio.ToolbarItems.Item.md)
- [Parent](Visio.ToolbarItems.Parent.md)
- [ParentItem](Visio.ToolbarItems.ParentItem.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]