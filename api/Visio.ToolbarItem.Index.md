---
title: ToolbarItem.Index property (Visio)
keywords: vis_sdr.chm13513695
f1_keywords:
- vis_sdr.chm13513695
ms.prod: visio
api_name:
- Visio.ToolbarItem.Index
ms.assetid: 2f97fbee-595b-4b71-137b-55de5a469ae6
ms.date: 06/08/2017
localization_priority: Normal
---


# ToolbarItem.Index property (Visio)

Gets the ordinal position of a  **ToolbarItem** object in the **ToolbarItems** collection. Read-only.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[ToolbarItem](Visio.ToolbarItem.md)** object.


## Return value

Long


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

These collections are indexed starting with 0:  **AccelItems**, **AccelTables**, **MenuSets**, **MenuItems**, **Menus**, **ToolbarItems**, **Toolbars**, and **ToolbarSets**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]