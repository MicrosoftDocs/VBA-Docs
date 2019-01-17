---
title: CommandBar.ShowPopup method (Office)
keywords: vbaof11.chm3017
f1_keywords:
- vbaof11.chm3017
ms.prod: office
api_name:
- Office.CommandBar.ShowPopup
ms.assetid: e501b7d2-2606-976c-b391-1aa8fa07f105
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.ShowPopup method (Office)

Displays a command bar as a shortcut menu at the specified coordinates or at the current pointer coordinates.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**ShowPopup**(_x_, _y_)

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _x_|Optional|**Variant**|The _x_-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current _x_-coordinate of the pointer is used.|
| _y_|Optional|**Variant**|The _y_-coordinate on which the location of the shortcut menu is based. If this argument is omitted, the current _y_-coordinate of the pointer is used.|

## Remarks

When menus are left-aligned, the shortcut menu displayed by the **ShowPopup** method has its upper left corner at (_x_, _y_ + 1); when menus are right-aligned, the shortcut menu has its upper right corner at (_x_ + 1, _y_ + 1). You can use the Windows function **GetSystemMetrics(SM_MENUDROPALIGNMENT)** to check the system metric for dropdown menu alignment.

When the screen location of the (_x_, _y_) coordinates would cause all or part of the popup menu to be displayed beyond the edge of the visible screen, the popup menu shifts to fit into the viewable area.


## Example

This example creates a shortcut menu containing two controls. The **ShowPopup** method is used to make the shortcut menu visible.


```vb
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarPopup, Temporary:=False) 
With myBar 
    .Controls.Add Type:=msoControlButton, Id:=3 
    .Controls.Add Type:=msoControlComboBox 
End With 
myBar.ShowPopup
```

> [!NOTE]
> If the **Position** property of the command bar is not set to **msoBarPopup**, this method fails.


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]