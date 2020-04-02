---
title: Menus object (Visio)
keywords: vis_sdr.chm10165
f1_keywords:
- vis_sdr.chm10165
ms.prod: visio
api_name:
- Visio.Menus
ms.assetid: 0c487176-1857-d496-8b2e-6a6aae668c6f
ms.date: 06/19/2019
localization_priority: Normal
---


# Menus object (Visio)

Includes a **[Menu](Visio.Menu.md)** object for each menu in a Microsoft Visio menu set.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **Menus** collection, use the **[Menus](visio.menuset.menus.md)** property of a **MenuSet** object.

The default property of a **Menus** collection is **Item**.

Unlike other Visio collections, the **Menus** collection is indexed starting with zero (0) rather than 1.

## Methods

- [Add](Visio.Menus.Add.md)
- [AddAt](Visio.Menus.AddAt.md)

## Properties

- [Count](Visio.Menus.Count.md)
- [Item](Visio.Menus.Item.md)
- [Parent](Visio.Menus.Parent.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]