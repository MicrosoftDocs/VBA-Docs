---
title: CommandBar.Name property (Office)
keywords: vbaof11.chm3010
f1_keywords:
- vbaof11.chm3010
ms.prod: office
api_name:
- Office.CommandBar.Name
ms.assetid: 4d578782-b59d-3dd7-be99-b9d79f8f3eaa
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Name property (Office)

Gets the name of the built-in **CommandBar** object. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

The local name of a built-in command bar is displayed in the title bar (when the command bar isn't docked) and in the list of available command bars—wherever that list is displayed in the container application. For a built-in command bar, the **Name** property returns the command bar's U.S. English name. Use the **NameLocal** property to return the localized name. If you change the value of the **LocalName** property for a custom command bar, the value of **Name** changes as well, and vice versa.


## Example

This example searches the collection of command bars for the command bar named **Custom**. If this command bar is found, the example makes it visible.


```vb
foundFlag =  False 
For Each bar In CommandBars 
    If bar.Name = "Custom" Then 
        foundFlag = True  
        bar.Visible = True  
    End If 
Next 
If Not foundFlag Then 
    MsgBox "'Custom' bar isn't in collection." 
Else 
    MsgBox "'Custom' bar is now visible." 
End If
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]