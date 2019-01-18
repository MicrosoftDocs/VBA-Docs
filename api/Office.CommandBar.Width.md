---
title: CommandBar.Width property (Office)
keywords: vbaof11.chm3021
f1_keywords:
- vbaof11.chm3021
ms.prod: office
api_name:
- Office.CommandBar.Width
ms.assetid: ae092193-59fd-25a1-c1d0-ebe6d6532756
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Width property (Office)

Gets or sets the width (in pixels) of the specified command bar. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Width**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Return value

Integer


## Example

This example adds a custom control to the command bar named **Custom**. The example sets the height of the custom control to twice the height of the command bar and sets its width to 50 pixels. Notice how the command bar automatically resizes itself to accommodate the control.


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