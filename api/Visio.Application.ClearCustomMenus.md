---
title: Application.ClearCustomMenus method (Visio)
keywords: vis_sdr.chm10016110
f1_keywords:
- vis_sdr.chm10016110
ms.prod: visio
api_name:
- Visio.Application.ClearCustomMenus
ms.assetid: 01c7f266-e940-b02c-b77d-7178c9296f98
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.ClearCustomMenus method (Visio)

Restores the built-in Microsoft Visio user interface.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**ClearCustomMenus**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


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