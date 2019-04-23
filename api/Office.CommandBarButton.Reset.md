---
title: CommandBarButton.Reset method (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Reset
ms.assetid: 0e39c960-3928-f91a-cf7e-1df5a2fd217b
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Reset method (Office)

Resets a built-in **CommandBarButton** control to its original function and face.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Reset**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state.


## Example

This example customizes a command bar button. First, the button properties are reset to their default state. Then various button properties are set. 


```vb
Dim cbButton As CommandBarButton 
Set cbButton = CommandBars("Custom").Controls(2) 
cbButton.Reset 
With cbButton 
    .BuiltInFace = True  
    .Caption = "Compute Total" 
    .DescriptionText = "This button computes the total of all purchases." 
    .Enabled = True  
    .TooltipText = "Click to compute total amount for all items in your cart." 
End With
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]