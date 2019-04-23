---
title: CommandBarButton.Height property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Height
ms.assetid: b374ae8b-cce2-7562-1247-32ea90dc3c68
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Height property (Office)

Gets or sets the height of a command bar control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

Integer


## Example

This example adds a custom control to the command bar named **Custom**. The example sets the height of the custom control to twice the height of the command bar, and sets the control's width to 50 pixels. Notice how the command bar automatically resizes itself to accommodate the control.


```vb
Set myBar = CommandBars("Custom") 
barHeight = myBar.Height 
Set myControl = myBar.Controls _ 
    .Add(Type:=msoControlButton, _ 
    Id:= CommandBars("Standard").Controls("Save").Id, _ 
     Temporary:=True) 
With myControl 
    .Height = barHeight * 2 
    .Width = 50 
End With 
myBar.Visible = True
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]