---
title: CommandBarComboBox.ListCount property (Office)
keywords: vbaof11.chm8006
f1_keywords:
- vbaof11.chm8006
ms.prod: office
api_name:
- Office.CommandBarComboBox.ListCount
ms.assetid: 3ab55501-b82e-0380-d805-e4386c399131
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.ListCount property (Office)

Gets the number of list items in a **CommandBarComboBox** control. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**ListCount**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Example

This example checks the number of items in the combo box on the command bar named **Custom**. If there aren't three items in the list that the procedure produces, the example displays a message advising the user that the combo box may be damaged and asks the user to reinstall the application.


```vb
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlComboBox, ID:=1 
    .Visible = True  
End With 
With CommandBars("Custom").Controls(1) 
    .AddItem "Get Stock Quote", 1 
    .AddItem "View Chart", 2 
    .AddItem "View Fundamentals", 3 
    .Caption = "Stock Data" 
    .DescriptionText = "View Data For Stock" 
End With 
If CommandBars("Custom").Controls(1).ListCount _ 
     > 4 Then 
MsgBox ("ComboBox appears to be damaged." & _ 
     " Please reinstall.") 
End If
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]