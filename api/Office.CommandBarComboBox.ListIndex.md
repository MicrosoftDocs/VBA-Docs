---
title: CommandBarComboBox.ListIndex property (Office)
keywords: vbaof11.chm8008
f1_keywords:
- vbaof11.chm8008
ms.prod: office
api_name:
- Office.CommandBarComboBox.ListIndex
ms.assetid: 3267a20a-7b33-3a89-5def-46c8b9756c04
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.ListIndex property (Office)

Gets or sets the index number of the selected item in the list portion of the **CommandBarComboBox** control. If nothing is selected in the list, this property returns zero. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**ListIndex**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Remarks

> [!NOTE]
> This property fails when applied to controls other than list controls.

Setting the **ListIndex** property causes the specified control to select the given item and execute the appropriate action in the application.


## Example

This example uses the **ListIndex** property to determine the correct subroutine to run, based on the selection in the combo box on the command bar named **My Custom Bar**. Because the procedure uses **ListIndex**, the text in the combo box can be anything.


```vb
Sub processSelection() 
Dim userChoice As Long 
userChoice = CommandBars("My Custom Bar").Controls(1).ListIndex 
    Select Case userChoice 
        Case 1 
            chartcourse 
        Case 2 
            displaygraph 
        Case Else 
            MsgBox ("Invalid choice. Please choose again.") 
    End Select 
End Sub
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]