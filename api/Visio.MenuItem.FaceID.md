---
title: MenuItem.FaceID property (Visio)
keywords: vis_sdr.chm12913495
f1_keywords:
- vis_sdr.chm12913495
ms.prod: visio
api_name:
- Visio.MenuItem.FaceID
ms.assetid: 1d4672d6-98e5-0875-4884-42f7d3ede52b
ms.date: 06/08/2017
localization_priority: Normal
---


# MenuItem.FaceID property (Visio)

Gets or sets the icon for an item. Read/write.


## Syntax

_expression_. `FaceID`

_expression_ A variable that represents a **[MenuItem](Visio.MenuItem.md)** object.


## Return value

Integer


## Remarks


> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use any of the constants prefixed with  **visIconIX** that are declared by the Visio type library in **[VisUIIconIDs](Visio.visuiiconids.md)**.

The  **FaceID** property determines a button's icon, but not its function. Use the **CmdNum** property of a **MenuItem** object to set a button's function.

The  **FaceID** property is the same as the **TypeSpecific1** property when the **CtrlType** property is type **visCtrlTypeBUTTON**, which is declared in the Visio type library in **[VisUICtrlTypes](Visio.visuictrltypes.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]