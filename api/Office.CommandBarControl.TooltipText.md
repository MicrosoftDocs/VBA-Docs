---
title: CommandBarControl.TooltipText property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.TooltipText
ms.assetid: 03e51dbd-0d5a-5094-545f-4a98a6508b4d
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.TooltipText property (Office)

Gets or sets the text displayed in the **ScreenTip** of a **CommandBarControl**. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**TooltipText**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Return value

String


## Remarks

By default, the value of the **Caption** property is used as the **ScreenTip**.


## Example

This example adds a **ScreenTip** to the last control on the active menu bar.


```vb
Set myMenuBar = CommandBars.ActiveMenuBar 
Set lastCtrl = myMenuBar _ 
   .Controls(myMenuBar.Controls.Count) 
lastCtrl.BeginGroup = True  
lastCtrl.TooltipText = "Click for help on UI feature"
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]