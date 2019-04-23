---
title: CommandBarPopup.Execute method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Execute
ms.assetid: fedebe76-86f5-9c30-6e23-a20e0024bbf4
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.Execute method (Office)

Runs the procedure or built-in command assigned to the specified **CommandBarPopup** control.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Execute**

_expression_ Required. A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Example

This Microsoft Excel example creates a command bar and then adds a built-in command bar button control to it. The button executes the Excel **AutoSum** function. This example uses the **Execute** method to total the selected range of cells when the command bar appears.


```vb
Dim cbrCustBar As CommandBar 
Dim ctlAutoSum As CommandBarButton 
Set cbrCustBar = CommandBars.Add("Custom") 
Set ctlAutoSum = cbrCustBar.Controls _ 
    .Add(msoControlButton, CommandBars("Standard") _ 
    .Controls("AutoSum").Id) 
cbrCustBar.Visible = True  
ctlAutoSum.Execute
```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]