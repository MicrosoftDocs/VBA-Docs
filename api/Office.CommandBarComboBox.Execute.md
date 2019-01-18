---
title: CommandBarComboBox.Execute method (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Execute
ms.assetid: 13ec7924-2420-c0c0-750f-4dae8b8e1503
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.Execute method (Office)

Runs the procedure or built-in command assigned to the specified **CommandBarComboBox** control.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Execute**

_expression_ Required. A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


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

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]