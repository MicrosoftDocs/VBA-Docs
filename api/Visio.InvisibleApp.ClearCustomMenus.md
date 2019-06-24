---
title: InvisibleApp.ClearCustomMenus method (Visio)
keywords: vis_sdr.chm17516110
f1_keywords:
- vis_sdr.chm17516110
ms.prod: visio
api_name:
- Visio.InvisibleApp.ClearCustomMenus
ms.assetid: bf075a03-4b3f-4f38-a6b4-6fbc2cc95fe1
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.ClearCustomMenus method (Visio)

Restores the built-in Microsoft Visio user interface.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**ClearCustomMenus**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Nothing


## Remarks

Calling the **ClearCustomMenus** method on an object without custom menus has no effect.


## Example

This example shows how to clear custom menus for the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** and **Application** objects and restore the built-in Visio menus.

```vb
 
Public Sub ClearCustomMenus_Example() 
 
 'Tell Visio to use the built-in menus. 
 ThisDocument.ClearCustomMenus 
 Visio.Application.ClearCustomMenus 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]