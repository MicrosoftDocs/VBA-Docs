---
title: InvisibleApp.ClearCustomToolbars method (Visio)
keywords: vis_sdr.chm17516115
f1_keywords:
- vis_sdr.chm17516115
ms.prod: visio
api_name:
- Visio.InvisibleApp.ClearCustomToolbars
ms.assetid: 3020ea80-ea8b-3670-865b-329326835a7f
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.ClearCustomToolbars method (Visio)

Restores the built-in Microsoft Visio user interface.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**ClearCustomToolbars**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Nothing


## Remarks

Calling the **ClearCustomToolbars** method on an object without custom toolbars has no effect.


## Example

This example shows how to clear custom toolbars for the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** and **Application** objects and restore the built-in Microsoft Visio toolbars.

```vb
 
Public Sub ClearCustomToolbars_Example() 
 
 'Tell Visio to use the built-in toolbars. 
 ThisDocument.ClearCustomToolbars 
 Visio.Application.ClearCustomToolbars 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]