---
title: CommandBarButton.OLEUsage property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.OLEUsage
ms.assetid: 4ff6f74d-4eed-8a30-468c-22be5dee1c7e
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.OLEUsage property (Office)

Gets or sets the OLE client and OLE server roles in which a **CommandBarButton** control will be used when two Microsoft Office applications are merged. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OLEUsage**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

msoControlOLEUsage


## Remarks

This property is intended to allow you to specify how individual add-in applications' command bar controls will be represented in one Office application when it is merged with another Office application. If both the client and server implement command bars, the command bar controls are embedded in the client control by control. Custom controls marked as client-only (or neither client nor server) are dropped from the server, and controls marked as server-only (or neither server nor client) are dropped from the client. The remaining controls are merged.

If one of the merging applications isn't an Office application, normal OLE menu merging is used, which is controlled by the **OLEMenuGroup** property.


## Example

This example adds a new button to the command bar named **Tools**, and sets its **OLEUsage** property.


```vb
Set myControl = CommandBars("Tools").Controls _ 
    .Add(Type:=msoControlButton,Temporary:=True) 
myControl.OLEUsage = msoControlOLEUsageNeither
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]