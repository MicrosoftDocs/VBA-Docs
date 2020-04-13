---
title: ToolbarItems.AddAt method (Visio)
keywords: vis_sdr.chm13616015
f1_keywords:
- vis_sdr.chm13616015
ms.prod: visio
api_name:
- Visio.ToolbarItems.AddAt
ms.assetid: 99611457-ada0-ce21-a3fe-fbc48fc88df9
ms.date: 06/08/2017
localization_priority: Normal
---


# ToolbarItems.AddAt method (Visio)

Creates a new **ToolbarItem** object at a specified index in the **ToolbarItems** collection. .


## Syntax

_expression_. `AddAt`( `_lIndex_` )

_expression_ A variable that represents a **[ToolbarItems](Visio.ToolbarItems.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|The index at which to add the object.|

## Return value

ToolbarItem


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the index is zero (0), the object is added at the beginning of the collection.

The beginning of a  **ToolbarItems** collection is the leftmost item in a toolbar that is arranged horizontally.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]