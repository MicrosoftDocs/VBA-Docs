---
title: ToolbarItems.Item property (Visio)
keywords: vis_sdr.chm13613765
f1_keywords:
- vis_sdr.chm13613765
ms.prod: visio
api_name:
- Visio.ToolbarItems.Item
ms.assetid: 0ef04285-aaaf-3bff-8758-2610fcd6d5f1
ms.date: 06/08/2017
localization_priority: Normal
---


# ToolbarItems.Item property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[ToolbarItems](Visio.ToolbarItems.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

ToolbarItem


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index )
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]