---
title: Application.CustomToolbarsFile property (Visio)
keywords: vis_sdr.chm10013360
f1_keywords:
- vis_sdr.chm10013360
ms.prod: visio
api_name:
- Visio.Application.CustomToolbarsFile
ms.assetid: e4759ee0-1128-8238-ad0b-47ad365ce88d
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.CustomToolbarsFile property (Visio)

Returns or sets the name of the file that defines custom toolbars and status bars for an **Application** object. Read/write.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**CustomToolbarsFile**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

String


## Remarks

If the object is not using custom toolbars, the **CustomToolbarsFile** property returns **Nothing**.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]