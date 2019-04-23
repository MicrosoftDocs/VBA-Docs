---
title: CommandBar.Context property (Office)
keywords: vbaof11.chm3002
f1_keywords:
- vbaof11.chm3002
ms.prod: office
api_name:
- Office.CommandBar.Context
ms.assetid: e7b8a7e5-0799-84e8-c7e3-5f713971099d
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Context property (Office)

Gets or sets a string that determines where a command bar will be saved. The string is defined and interpreted by the application. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Context**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

You can set the **Context** property only for custom command bars. This property will fail if the application doesn't recognize the context string, or if the application doesn't support changing context strings programmatically.


## Example

This example displays a message box containing the context string for the command bar named "Custom". This example works in Microsoft Word and other applications that support the **Context** property.


```vb
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlButton, ID:=2 
    .Visible = True  
End With 
MsgBox (myBar.Context) 

```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]