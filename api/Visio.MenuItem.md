---
title: MenuItem object (Visio)
keywords: vis_sdr.chm10155
f1_keywords:
- vis_sdr.chm10155
ms.prod: visio
api_name:
- Visio.MenuItem
ms.assetid: 7161bf25-fde8-09d6-0c10-52a65f86feba
ms.date: 06/19/2019
localization_priority: Normal
---


# MenuItem object (Visio)

Represents a single menu item on a Microsoft Visio menu.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

The default property of **MenuItem** is **Caption**.

A **MenuItem** object contains all the information it needs to display the menu item and launch the appropriate Visio command or add-on. It also contains text for the **Undo**, **Redo**, and **Repeat** menu items and error messages.

The index of a **MenuItem** object within the **[MenuItems](Visio.MenuItems.md)** collection corresponds to the menu item's position from top to bottom on the menu or submenu, starting with zero (0).

If the menu item displays a submenu, the **MenuItem** object has a **MenuItems** collection that represents items on the submenu. The **MenuItem** object's **Caption** property contains the submenu title. Most of the other properties of the **MenuItem** object are ignored, because this object serves much the same role as a **Menu** object.

## Methods

-  [Delete](Visio.MenuItem.Delete.md)
-  [IconFileName](Visio.MenuItem.IconFileName.md)

## Properties

-  [ActionText](Visio.MenuItem.ActionText.md)
-  [AddOnArgs](Visio.MenuItem.AddOnArgs.md)
-  [AddOnName](Visio.MenuItem.AddOnName.md)
-  [BeginGroup](Visio.MenuItem.BeginGroup.md)
-  [BuiltIn](Visio.MenuItem.BuiltIn.md)
-  [Caption](Visio.MenuItem.Caption.md)
-  [CmdNum](Visio.MenuItem.CmdNum.md)
-  [CntrlType](Visio.MenuItem.CntrlType.md)
-  [Enabled](Visio.MenuItem.Enabled.md)
-  [FaceID](Visio.MenuItem.FaceID.md)
-  [Index](Visio.MenuItem.Index.md)
-  [IsHierarchical](Visio.MenuItem.IsHierarchical.md)
-  [MenuItems](Visio.MenuItem.MenuItems.md)
-  [PaletteWidth](Visio.MenuItem.PaletteWidth.md)
-  [Parent](Visio.MenuItem.Parent.md)
-  [State](Visio.MenuItem.State.md)
-  [Style](Visio.MenuItem.Style.md)
-  [TypeSpecific1](Visio.MenuItem.TypeSpecific1.md)
-  [TypeSpecific2](Visio.MenuItem.TypeSpecific2.md)
-  [Visible](Visio.MenuItem.Visible.md)
-  [Width](Visio.MenuItem.Width.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]