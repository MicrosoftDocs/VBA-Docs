---
title: CommandBarComboBox.Id property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Id
ms.assetid: 9cc143cb-4063-b397-05c9-d50a7c2efcb0
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.Id property (Office)

Gets the ID for a built-in **CommandBarComboBox** control. Read-only.


## Syntax

_expression_.**Id**

_expression_ Required. A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Remarks

A control's ID determines the built-in action for that control. The value of the **Id** property for all custom controls is 1.


## Example

This example changes the button face of the first control on the command bar named **Custom2** if the button's **ID** value is less than 25.

```vb
Set ctrl = CommandBars("Custom").Controls(1) 
With ctrl 
 If .Id < 25 Then 
 .FaceId = 17 
 .Tag = "Changed control" 
 End If 
End With
```

<br/>

The following example changes the caption of every control on the toolbar named **Standard** to the current value of the **Id** property for that control.

```vb
For Each ctl In CommandBars("Standard").Controls 
 ctl.Caption = CStr(ctl.Id) 
Next ctl
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]