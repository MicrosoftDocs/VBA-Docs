---
title: CommandBarComboBox.List property (Office)
keywords: vbaof11.chm8005
f1_keywords:
- vbaof11.chm8005
ms.prod: office
api_name:
- Office.CommandBarComboBox.List
ms.assetid: c90fae92-daab-1b08-6e85-8caae26d0b72
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.List property (Office)

Gets or sets an item in the **CommandBarComboBox** control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**List** (_Index_)

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Integer**| The list item to be set.|

## Remarks

This property is read-only for built-in combo box controls.


## Example

This example checks the fourth list item in the combo box control whose caption is **Stock Data** on the command bar named **Custom**. If the item isn't "View News," the example displays a message advising the user that the combo box may be damaged and asks the user to reinstall the application.


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
    .AddItem "View News", 4 
    .Caption = "Stock Data" 
    .DescriptionText = "View Data For Stock" 
End With 
If CommandBars("Custom").Controls(1).List(4) _ 
     > "View News" Then 
MsgBox ("Stock Data appears to be damaged." & _ 
     " Please reinstall application.") 
End If
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]