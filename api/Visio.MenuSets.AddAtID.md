---
title: MenuSets.AddAtID method (Visio)
keywords: vis_sdr.chm13416020
f1_keywords:
- vis_sdr.chm13416020
ms.prod: visio
api_name:
- Visio.MenuSets.AddAtID
ms.assetid: fab8fb43-961d-e184-0904-b57a036e62ca
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuSets.AddAtID method (Visio)

Creates a new object with a specified ID in a collection.


## Syntax

_expression_.**AddAtID** (_lID_)

_expression_ A variable that represents a **[MenuSets](Visio.MenuSets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lID_|Required| **Long**| The window context for the new object.|

## Return value

MenuSet


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The ID corresponds to a window or context menu. If the collection already contains an object at the specified ID, the  **AddAtID** method returns an error.

Valid IDs are declared by the Visio type library in member  **[VisUIObjSets](Visio.visuiobjsets.md)**. Not all collections include an object for every possible ID.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]