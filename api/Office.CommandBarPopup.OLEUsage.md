---
title: CommandBarPopup.OLEUsage property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.OLEUsage
ms.assetid: 75d338e0-f5ca-f4b6-2f94-e575749e6ae9
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.OLEUsage property (Office)

Gets or sets the OLE client and OLE server roles in which a **CommandBarPopup** control is used when two Microsoft Office applications are merged. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OLEUsage**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

MsoControlOLEUsage


## Remarks

This property is intended to allow you to specify how individual add-in applications' command bar controls are represented in one Office application when it is merged with another Office application. If both the client and server implement command bars, the command bar controls are embedded in the client control by control. Custom controls marked as client-only (or neither client nor server) are dropped from the server, and controls marked as server-only (or neither server nor client) are dropped from the client. The remaining controls are merged.

If one of the merging applications is not an Office application, normal OLE menu merging is used, which is controlled by the **OLEMenuGroup** property.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]