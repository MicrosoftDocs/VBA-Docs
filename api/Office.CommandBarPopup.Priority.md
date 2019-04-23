---
title: CommandBarPopup.Priority property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Priority
ms.assetid: cef115fd-fdc8-d8a3-b51d-c9fbc21a810f
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.Priority property (Office)

Gets or sets the priority of a **CommandBarPopup** control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Priority**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

Integer


## Remarks

A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left.


## Example

The following example sets the descriptive text and priority of a command bar popup.


```vb
Dim popControl As CommandBarPopup 
Set popControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics") 
 
With popControl. 
            .DescriptionText = "Graphics Selection dialog" 
            .Priority = 5 
End With 

```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]