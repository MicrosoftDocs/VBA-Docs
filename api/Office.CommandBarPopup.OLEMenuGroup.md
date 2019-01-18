---
title: CommandBarPopup.OLEMenuGroup property (Office)
keywords: vbaof11.chm7003
f1_keywords:
- vbaof11.chm7003
ms.prod: office
api_name:
- Office.CommandBarPopup.OLEMenuGroup
ms.assetid: 32b1bc39-19bc-d0ed-59b5-2e7fa03f329e
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.OLEMenuGroup property (Office)

Gets or sets an **[MsoOLEMenuGroup](office.msoolemenugroup.md)** constant that represents the menu group that the specified **CommandBarPopup** control belongs to when the menu groups of the OLE server are merged with the menu groups of an OLE client (that is, when an object of the container application type is embedded in another application). Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OLEMenuGroup**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Remarks

> [!NOTE] 
> This property is read-only for built-in controls.

This property is intended to allow add-in applications to specify how their command bar controls will be represented in the Office application. If either the container or the server does not implement command bars, normal OLE menu merging occurs: the menu bar is merged, and all the toolbars from the server, and none of the toolbars from the container. This property is relevant only for pop-up controls on the menu bar because menus are merged on the basis of their menu group category.

If both of the merging applications implement command bars, command bar controls are merged according to the **OLEUsage** property.


## Example

This example checks the **OLEMenuGroup** property of a new custom popup control on the command bar named **Custom** and sets the property to **msoOLEMenuGroupNone**.


```vb
Set myControl = CommandBars("Custom").Controls _ 
    .Add(Type:=msoControlPopup,Temporary:=False) 
myControl.OLEMenuGroup = msoOLEMenuGroupNone
```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]