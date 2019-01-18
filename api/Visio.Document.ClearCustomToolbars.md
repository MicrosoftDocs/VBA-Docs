---
title: Document.ClearCustomToolbars Method (Visio)
keywords: vis_sdr.chm10516115
f1_keywords:
- vis_sdr.chm10516115
ms.prod: visio
api_name:
- Visio.Document.ClearCustomToolbars
ms.assetid: 823877b1-ee82-f87e-d68f-d8c6010457cc
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ClearCustomToolbars Method (Visio)

Restores the built-in Microsoft Visio user interface.


## Syntax

 _expression_. `ClearCustomToolbars`

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


## Return value

Nothing


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Calling the  **ClearCustomToolbars** method on an object without custom toolbars has no effect.


## Example

This example shows how to clear custom toolbars for the  **ThisDocument** and **Application** objects and restore the built-in Microsoft Visio toolbars.


```vb
 
Public Sub ClearCustomToolbars_Example() 
 
 'Tell Visio to use the built-in toolbars. 
 ThisDocument.ClearCustomToolbars 
 Visio.Application.ClearCustomToolbars 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]