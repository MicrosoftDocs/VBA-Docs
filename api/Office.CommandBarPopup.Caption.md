---
title: CommandBarPopup.Caption property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Caption
ms.assetid: fc9221e6-cfb0-9f2a-290b-73a434569e65
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.Caption property (Office)

Gets or sets the caption text for a command bar control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Caption**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

String


## Example

This example adds a command bar control with a spelling checker button face to a custom command bar, and then it sets the caption to **Spelling checker**.


```vb
Set myBar = CommandBars.Add(Name:="Custom", _ 
Position:=msoBarTop, Temporary:=True) 
myBar.Visible = True  
Set myControl = myBar.Controls _ 
.Add(Type:=msoControlButton, Id:=2) 
With myControl 
    .DescriptionText = "Starts the spelling checker" 
    .Caption = "Spelling checker" 
End With
```

> [!NOTE] 
> A control's caption is also displayed as its default **ScreenTip**.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]