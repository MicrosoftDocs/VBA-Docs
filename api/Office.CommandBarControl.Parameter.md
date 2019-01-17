---
title: CommandBarControl.Parameter property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Parameter
ms.assetid: 6a1fd988-0c3f-3945-307f-e4e647c3642c
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.Parameter property (Office)

Gets or sets a string that an application can use to execute a command from a **CommandBarControl**. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Parameter**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Return value

String


## Remarks

If the specified parameter is set for a built-in control, the application can modify its default behavior if it can parse and use the new value. If the parameter is set for custom controls, it can be used to send information to Visual Basic procedures, or it can be used to hold information about the control (similar to a second **Tag** property value).


## Example

This example assigns a new parameter to a control and sets the focus to the new button.


```vb
Set myControl = CommandBars("Custom").Controls(4) 
With myControl 
    .Copy , 1 
    .Parameter = "2" 
    .SetFocus 
End With
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]