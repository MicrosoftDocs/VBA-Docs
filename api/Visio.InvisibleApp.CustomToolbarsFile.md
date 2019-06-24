---
title: InvisibleApp.CustomToolbarsFile property (Visio)
keywords: vis_sdr.chm17513360
f1_keywords:
- vis_sdr.chm17513360
ms.prod: visio
api_name:
- Visio.InvisibleApp.CustomToolbarsFile
ms.assetid: 0874023f-1e61-7842-be7d-9abe5c4ec63c
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.CustomToolbarsFile property (Visio)

Returns or sets the name of the file that defines custom toolbars and status bars for an **InvisibleApp** object. Read/write.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**CustomToolbarsFile**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

String


## Remarks

If the object is not using custom toolbars, the **CustomToolbarsFile** property returns **Nothing**.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]