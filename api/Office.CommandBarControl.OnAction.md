---
title: CommandBarControl.OnAction property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.OnAction
ms.assetid: 05e40fcb-ff67-049f-6386-a9ef20b48c87
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.OnAction property (Office)

Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a **CommandBarControl**. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OnAction**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Return value

String


## Remarks

The container application determines whether the value is a valid macro name.


## Example

This example adds a command bar control to the command bar named **Custom**. The procedure named **MySub** will run each time the control is clicked.

```vb
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls _ 
    .Add(Type:=msocontrolButton) 
With myControl 
    .FaceId = 2 
    .OnAction = "MySub" 
End With 
myBar.Visible = True
```

<br/>

This example adds a command bar control to the command bar named **Custom**. The COM add-in named **FinanceAddIn** will run each time the control is clicked.

```vb
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls _ 
    .Add(Type:=msocontrolButton) 
With myControl 
    .FaceId = 2 
    .OnAction = "!<FinanceAddIn>" 
End With 
myBar.Visible = True
```


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]