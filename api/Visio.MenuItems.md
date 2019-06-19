---
title: MenuItems object (Visio)
keywords: vis_sdr.chm10160
f1_keywords:
- vis_sdr.chm10160
ms.prod: visio
api_name:
- Visio.MenuItems
ms.assetid: 7799eff9-5432-9c44-2e74-345479eef5b6
ms.date: 06/19/2019
localization_priority: Normal
---


# MenuItems object (Visio)

Contains a **[MenuItem](Visio.MenuItem.md)** object for each command on a Microsoft Visio menu.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **MenuItems** collection, use the **MenuItems** property of a **Menu** object or a **MenuItem** object.

The default property of a **MenuItems** collection is **Item**.

Unlike other Visio collections, the **MenuItems** collection is indexed starting with zero (0) rather than 1.

## Methods

-  [Add](Visio.MenuItems.Add.md)
-  [AddAt](Visio.MenuItems.AddAt.md)

## Properties

-  [Count](Visio.MenuItems.Count.md)
-  [Item](Visio.MenuItems.Item.md)
-  [Parent](Visio.MenuItems.Parent.md)
-  [ParentItem](Visio.MenuItems.ParentItem.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]