---
title: CommandBarButton.Id property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Id
ms.assetid: d559a98c-b9b2-a987-c7af-278734a9545d
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Id property (Office)

Gets the ID for a built-in **CommandBarButton** control. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Id**

_expression_ Required. A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

A control's ID determines the built-in action for that control. The value of the **Id** property for all custom controls is 1.


## Example

This example changes the button face of the first control on the command bar named **Custom2** if the button's ID value is less than 25.


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

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]