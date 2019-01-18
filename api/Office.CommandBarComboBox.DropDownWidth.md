---
title: CommandBarComboBox.DropDownWidth property (Office)
keywords: vbaof11.chm8004
f1_keywords:
- vbaof11.chm8004
ms.prod: office
api_name:
- Office.CommandBarComboBox.DropDownWidth
ms.assetid: 051ac285-c7f1-a2b7-0c9a-ed2cb08cadc9
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.DropDownWidth property (Office)

Gets or sets the width (in pixels) of the list for the specified command bar combo box control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**DropDownWidth**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Remarks

> [!NOTE]
> An error occurs if you attempt to set this property for a built-in control.

If this property is set to -1, the width of the list is based on the length of the longest item in the combo box list. If this property is set to 0, the width of the list is based on the width of the control.


## Example

This example adds a combo box control containing two items to the command bar named **Custom**. The example also sets the number of line items, the width of the combo box, and an empty default for the combo box.


```vb
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls.Add(Type:=msoControlComboBox, Id:=1) 
With myControl 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListHeaderCount = 0 
End With
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]