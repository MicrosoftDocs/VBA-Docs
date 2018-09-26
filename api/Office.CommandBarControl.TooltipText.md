---
title: CommandBarControl.TooltipText Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.TooltipText
ms.assetid: 03e51dbd-0d5a-5094-545f-4a98a6508b4d
ms.date: 06/08/2017
---


# CommandBarControl.TooltipText Property (Office)

Gets or sets the text displayed in a  **CommandBarControl's** **ScreenTip**. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `TooltipText`

 _expression_ A variable that represents a [CommandBarControl](./Office.CommandBarControl.md) object.


### Return value

String


## Remarks

By default, the value of the  **Caption** property is used as the **ScreenTip**.


## Example

This example adds a  **ScreenTip** to the last control on the active menu bar.


```vb
Set myMenuBar = CommandBars.ActiveMenuBar 
Set lastCtrl = myMenuBar _ 
   .Controls(myMenuBar.Controls.Count) 
lastCtrl.BeginGroup = True  
lastCtrl.TooltipText = "Click for help on UI feature"
```


## See also


[CommandBarControl Object](Office.CommandBarControl.md)



[CommandBarControl Object Members](./overview/Library-Reference/commandbarcontrol-members-office.md)

