---
title: MenuItems.Item property (Visio)
keywords: vis_sdr.chm13013765
f1_keywords:
- vis_sdr.chm13013765
ms.prod: visio
api_name:
- Visio.MenuItems.Item
ms.assetid: 1324184f-5eee-460a-e0a9-7d714a8f561c
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuItems.Item property (Visio)

Returns a **MenuItem** object from the **MenuItems** collection. Read-only.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[MenuItems](Visio.MenuItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|


## Return value

MenuItem


## Remarks

When retrieving objects from a collection, you can omit **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:

```vb
objRet = object(index )
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]