---
title: CommandBarComboBox.SetFocus method (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.SetFocus
ms.assetid: 3170651c-40da-5025-8b36-195b836c8fcb
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.SetFocus method (Office)

Moves the keyboard focus to the specified **CommandBarComboBox** control. If the control is disabled or isn't visible, this method will fail.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**SetFocus**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


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

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]