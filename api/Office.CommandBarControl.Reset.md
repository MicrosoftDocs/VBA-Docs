---
title: CommandBarControl.Reset method (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Reset
ms.assetid: 7b2d42c4-ac1c-209e-6fe8-bd5ec91d1c57
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.Reset method (Office)

Resets a built-in command bar to its default configuration, or resets a built-in **CommandBarControl** to its original function and face.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Reset**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state. Resetting a built-in command bar removes custom controls and restores built-in controls.


## Example

This example uses the value of user to adjust the command bars according to the user level. If user is "Level 1," the command bar named **Custom** is displayed. If user is any other value, the built-in Visual Basic command bar is reset to its default state and the command bar named **Custom** is disabled.


```vb
Set myBarControl = CommandBars("Custom").Controls(2) 
If user = "Level 1" Then 
    myBarControl.Visible = True  
Else 
    CommandBars("Visual Basic").Reset 
    myBarControl.Enabled = False  
End If
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]