---
title: Menu.CmdNum property (Visio)
keywords: vis_sdr.chm13113255
f1_keywords:
- vis_sdr.chm13113255
ms.prod: visio
api_name:
- Visio.Menu.CmdNum
ms.assetid: 13754873-94bd-3497-829c-374aec3615da
ms.date: 06/08/2017
localization_priority: Normal
---


# Menu.CmdNum property (Visio)

Gets or sets the command ID associated with a menu. Read/write.


## Syntax

_expression_.**CmdNum**

_expression_ A variable that represents a **[Menu](Visio.Menu.md)** object.


## Return value

Integer


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Valid command IDs are declared by the Visio type library in  **[VisUICmds](Visio.visuicmds.md)**. They have the prefix **visCmd**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]