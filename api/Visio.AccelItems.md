---
title: AccelItems object (Visio)
keywords: vis_sdr.chm10015
f1_keywords:
- vis_sdr.chm10015
ms.prod: visio
api_name:
- Visio.AccelItems
ms.assetid: 0ea77c63-1fe4-4edf-0b7b-2293eb4ed180
ms.date: 06/19/2019
localization_priority: Normal
---


# AccelItems object (Visio)

Includes an **[AccelItem](Visio.AccelItem.md)** object for each accelerator in a Microsoft Visio window context.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve an **AccelItems** collection, use the **[AccelItems](visio.acceltable.accelitems.md)** property of an **AccelTable** object.

The default property for an **AccelItems** collection is **Item**.

Unlike other Visio collections, the **AccelItems** collection is indexed starting with zero (0) rather than 1.

## Methods

- [Add](Visio.AccelItems.Add.md)

## Properties

- [Count](Visio.AccelItems.Count.md)
- [Item](Visio.AccelItems.Item.md)
- [Parent](Visio.AccelItems.Parent.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]