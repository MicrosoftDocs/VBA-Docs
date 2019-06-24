---
title: AccelTables.AddAtID method (Visio)
keywords: vis_sdr.chm14816020
f1_keywords:
- vis_sdr.chm14816020
ms.prod: visio
api_name:
- Visio.AccelTables.AddAtID
ms.assetid: 581526c5-eebb-f79a-e48c-b716be719c6f
ms.date: 06/24/2019
localization_priority: Normal
---


# AccelTables.AddAtID method (Visio)

Creates a new object with a specified ID in a collection.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**AddAtID** (_lID_)

_expression_ A variable that represents an **[AccelTables](Visio.AccelTables.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lID_|Required| **Long**| The window context for the new object.|

## Return value

**[AccelTable](Visio.AccelTable.md)**


## Remarks

The ID corresponds to a window or context menu. If the collection already contains an object at the specified ID, the **AddAtID** method returns an error.

Valid IDs are declared by the Visio type library in member **[VisUIObjSets](Visio.visuiobjsets.md)**. Not all collections include an object for every possible ID.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]