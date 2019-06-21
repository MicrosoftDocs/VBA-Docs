---
title: MenuSets.Item property (Visio)
keywords: vis_sdr.chm13413765
f1_keywords:
- vis_sdr.chm13413765
ms.prod: visio
api_name:
- Visio.MenuSets.Item
ms.assetid: a7ad3a73-33ec-1e69-c6d6-7356876be53c
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuSets.Item property (Visio)

Returns a  **MenuSet** object from the **MenuSets** collection. Read-only.


## Syntax

_expression_.**Item** (_lIndex_)

_expression_ A variable that represents a **[MenuSets](Visio.MenuSets.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

MenuSet


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index )
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]