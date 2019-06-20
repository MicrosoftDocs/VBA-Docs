---
title: AccelTables object (Visio)
keywords: vis_sdr.chm10025
f1_keywords:
- vis_sdr.chm10025
ms.prod: visio
api_name:
- Visio.AccelTables
ms.assetid: 1bc9671b-83dc-1349-9171-92d1650ebec8
ms.date: 06/19/2019
localization_priority: Normal
---


# AccelTables object (Visio)

Includes an **[AccelTable](Visio.AccelTable.md)** object for each Microsoft Visio window context that has accelerators.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve an **AccelTables** collection, use the **[AccelTables](visio.uiobject.acceltables.md)** property of a **UIObject** object.

The default property of **AccelTables** is **Item**.

Unlike other Visio collections, the **AccelTables** collection is indexed starting with zero (0) rather than 1.

An **AccelTable** object is identified in the **AccelTables** collection by its **[SetID](Visio.AccelTable.SetID.md)** property, which corresponds to a Visio window context. For a list of **SetID** values that identify **AccelTable** objects, see the **SetID** property.

## Methods

-  [Add](Visio.AccelTables.Add.md)
-  [AddAtID](Visio.AccelTables.AddAtID.md)

## Properties

-  [Count](Visio.AccelTables.Count.md)
-  [Item](Visio.AccelTables.Item.md)
-  [ItemAtID](Visio.AccelTables.ItemAtID.md)
-  [Parent](Visio.AccelTables.Parent.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]