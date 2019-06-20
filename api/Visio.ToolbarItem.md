---
title: ToolbarItem object (Visio)
keywords: vis_sdr.chm10275
f1_keywords:
- vis_sdr.chm10275
ms.prod: visio
api_name:
- Visio.ToolbarItem
ms.assetid: 2f0798cf-f31e-e213-d9db-325d58a77e96
ms.date: 06/19/2019
localization_priority: Normal
---


# ToolbarItem object (Visio)

Represents one item in a **[Toolbar](Visio.Toolbar.md)** object. A **ToolbarItem** object can represent a button, combo box, or any other item on the Microsoft Visio toolbars.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

The index of the **ToolbarItem** object within the **[ToolbarItems](Visio.ToolbarItems.md)** collection corresponds to its position on the toolbar, starting with zero (0) for the item farthest to the left if the toolbars are arranged horizontally.

Beginning with Microsoft Visio 2002, use the **BeginGroup** property to create spaces on a toolbar.

## Methods

-  [Delete](Visio.ToolbarItem.Delete.md)
-  [IconFileName](Visio.ToolbarItem.IconFileName.md)

## Properties

-  [ActionText](Visio.ToolbarItem.ActionText.md)
-  [AddOnArgs](Visio.ToolbarItem.AddOnArgs.md)
-  [AddOnName](Visio.ToolbarItem.AddOnName.md)
-  [BeginGroup](Visio.ToolbarItem.BeginGroup.md)
-  [BuiltIn](Visio.ToolbarItem.BuiltIn.md)
-  [Caption](Visio.ToolbarItem.Caption.md)
-  [CmdNum](Visio.ToolbarItem.CmdNum.md)
-  [CntrlType](Visio.ToolbarItem.CntrlType.md)
-  [Enabled](Visio.ToolbarItem.Enabled.md)
-  [FaceID](Visio.ToolbarItem.FaceID.md)
-  [Index](Visio.ToolbarItem.Index.md)
-  [IsHierarchical](Visio.ToolbarItem.IsHierarchical.md)
-  [PaletteWidth](Visio.ToolbarItem.PaletteWidth.md)
-  [Parent](Visio.ToolbarItem.Parent.md)
-  [State](Visio.ToolbarItem.State.md)
-  [Style](Visio.ToolbarItem.Style.md)
-  [ToolbarItems](Visio.ToolbarItem.ToolbarItems.md)
-  [TypeSpecific1](Visio.ToolbarItem.TypeSpecific1.md)
-  [TypeSpecific2](Visio.ToolbarItem.TypeSpecific2.md)
-  [Visible](Visio.ToolbarItem.Visible.md)
-  [Width](Visio.ToolbarItem.Width.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]