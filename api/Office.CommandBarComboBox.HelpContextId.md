---
title: CommandBarComboBox.HelpContextId property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.HelpContextId
ms.assetid: 3b34572b-af1b-a4fc-a98e-23d51315a077
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.HelpContextId property (Office)

Gets or sets the Help context Id number for the Help topic attached to the **CommandBarComboBox** control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HelpContextId**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Return value

Integer


## Remarks

To use this property, you must also set the **HelpFile** property. Help topics respond to Shift+F1.


## Example

This example adds a custom command bar with a combo box that tracks stock data. The example also specifies the Help topic to be displayed for the combo box when the user presses Shift+F1.


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
    .HelpFile = "C:\corphelp\custom.hlp" 
    .HelpContextID = 47 
End With
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]