---
title: MenuItem.CmdNum property (Visio)
keywords: vis_sdr.chm12913255
f1_keywords:
- vis_sdr.chm12913255
ms.prod: visio
api_name:
- Visio.MenuItem.CmdNum
ms.assetid: 7902ad54-62e3-f8da-ea34-7af43f2f13ef
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuItem.CmdNum property (Visio)

Gets or sets the command ID associated with a menu item. Read/write.


## Syntax

_expression_.**CmdNum**

_expression_ A variable that represents a **[MenuItem](Visio.MenuItem.md)** object.


## Return value

Integer


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When the  **AddOnName** property of a **MenuItem** object indicates an add-on to run, Microsoft Visio automatically assigns a **CmdNum** property.

To insert a separator in a menu preceding a  **MenuItem** object, use the **BeginGroup** property.

Valid command IDs are declared by the Visio type library in  **[VisUICmds](Visio.visuicmds.md)**. They have the prefix **visCmd**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]