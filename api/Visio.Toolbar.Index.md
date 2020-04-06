---
title: Toolbar.Index property (Visio)
keywords: vis_sdr.chm13713695
f1_keywords:
- vis_sdr.chm13713695
ms.prod: visio
api_name:
- Visio.Toolbar.Index
ms.assetid: 8af96f5a-1c41-633c-3542-d720712444bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Toolbar.Index property (Visio)

Gets the ordinal position of a  **Toolbar** object in a **Toolbars** collection. Read-only.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[Toolbar](Visio.Toolbar.md)** object.


## Return value

Long


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

These collections are indexed starting with 0:  **AccelItems**, **AccelTables**, **MenuSets**, **MenuItems**, **Menus**, **ToolbarItems**, **Toolbars**, and **ToolbarSets**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]