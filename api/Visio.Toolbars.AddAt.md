---
title: Toolbars.AddAt method (Visio)
keywords: vis_sdr.chm13816015
f1_keywords:
- vis_sdr.chm13816015
ms.prod: visio
api_name:
- Visio.Toolbars.AddAt
ms.assetid: 925f6c3a-8d74-9359-4008-0fced3e03ec1
ms.date: 06/08/2017
localization_priority: Normal
---


# Toolbars.AddAt method (Visio)

Creates a new **Toolbar** object at a specified index in the **Toolbars** collection. .


## Syntax

_expression_. `AddAt`( `_lIndex_` )

_expression_ A variable that represents a **[Toolbars](Visio.Toolbars.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|The index at which to add the object.|

## Return value

Toolbar


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If the index is zero (0), the object is added at the beginning of the collection.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]