---
title: CommandBarControl.Id property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Id
ms.assetid: 0931a07a-4a6b-cc84-a43b-b57ea9a22b78
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.Id property (Office)

Gets the ID for a built-in **CommandBarControl**. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Id**

_expression_ Required. A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Remarks

A control's ID determines the built-in action for that control. The value of the **Id** property for all custom controls is 1.


## Example

This example changes the button face of the first control on the command bar named **Custom2** if the button's **Id** value is less than 25.

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

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]