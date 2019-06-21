---
title: MenuSets.ItemAtID property (Visio)
keywords: vis_sdr.chm13413770
f1_keywords:
- vis_sdr.chm13413770
ms.prod: visio
api_name:
- Visio.MenuSets.ItemAtID
ms.assetid: d05dce0a-c01e-d249-a88d-44d246404ee0
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuSets.ItemAtID property (Visio)

Returns the  **MenuSet** object with the specified ID within the **MenuSets** collection. Read-only.


## Syntax

_expression_. `ItemAtID`( `_lID_` )

_expression_ A variable that represents a **[MenuSets](Visio.MenuSets.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lID_|Required| **Long**|The Visio context ID of the object to retrieve.|

## Return value

MenuSet


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The ID corresponds to a window or context menu. Constants for IDs are prefixed with  **visUIObjectSet** and are declared by the Visio type library in **[VisUIObjSets](Visio.visuiobjsets.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]