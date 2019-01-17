---
title: CommandBarComboBox.DropDownLines property (Office)
keywords: vbaof11.chm8003
f1_keywords:
- vbaof11.chm8003
ms.prod: office
api_name:
- Office.CommandBarComboBox.DropDownLines
ms.assetid: 715bbec9-1bd6-c7b0-0d1e-e57d61689d52
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.DropDownLines property (Office)

Gets or sets the number of lines in a command bar combo box control. The combo box control must be a custom control and it must be a drop-down list box or a combo box. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**DropDownLines**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Remarks

> [!NOTE]
> An error occurs if you attempt to set this property for a combo box control that's an edit box or a built-in combo box control.

If this property is set to 0 (zero), the number of lines in the control is based on the number of items in the list.


## Example

This example adds a combo box control containing two items to the command bar named **Custom**. The example also sets the number of line items, the width of the combo box, and an empty default for the combo box.


```vb
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls.Add(Type:=msoControlComboBox, Id:=1) 
With myControl 
    .AddItem Text:="First Item", Index:=1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListHeaderCount = 0 
End With
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]