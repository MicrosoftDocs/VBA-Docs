---
title: CommandBar.Height property (Office)
keywords: vbaof11.chm3007
f1_keywords:
- vbaof11.chm3007
ms.prod: office
api_name:
- Office.CommandBar.Height
ms.assetid: 9a5c84ae-29c0-0ff3-74f4-864c978336d2
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Height property (Office)

Gets or sets the height of a **CommandBar**. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Return value

Integer


## Example

This example adds a custom control to the command bar named **Custom**. The example sets the height of the custom control to twice the height of the command bar and sets the control's width to 50 pixels. Notice how the command bar automatically resizes itself to accommodate the control.


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

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]