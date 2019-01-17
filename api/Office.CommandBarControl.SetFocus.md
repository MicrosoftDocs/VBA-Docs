---
title: CommandBarControl.SetFocus method (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.SetFocus
ms.assetid: e20065eb-a1a3-f750-5585-6e38a328b946
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.SetFocus method (Office)

Moves the keyboard focus to the specified **CommandBarControl**. If the control is disabled or isn't visible, this method will fail.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**SetFocus**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Remarks

The focus on the control is subtle. After you use this method, you will notice a three-dimensional highlight on the control. Pressing the arrow keys will navigate in the toolbars, as if you had arrived at the control by pressing only keyboard controls.


## Example

This example creates a command bar named **Custom** and adds a **ComboBox** control and a **Button** control to it. The example then uses the **SetFocus** method to set the focus to the **ComboBox** control.


```vb
Set focusBar = CommandBars.Add(Name:="Custom") 
With CommandBars("Custom") 
    .Visible = True  
    .Position = msoBarTop 
End With 
 
Set testComboBox = CommandBars("Custom").Controls _ 
    .Add(Type:=msoControlComboBox, ID:=1) 
With testComboBox 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
End With 
Set testButton = CommandBars("Custom").Controls _ 
    .Add(Type:=msoControlButton) 
testButton.FaceId = 17 
' Set the focus to the combo box. 
testComboBox.SetFocus
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]