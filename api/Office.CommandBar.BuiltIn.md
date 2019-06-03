---
title: CommandBar.BuiltIn property (Office)
keywords: vbaof11.chm3001
f1_keywords:
- vbaof11.chm3001
ms.prod: office
api_name:
- Office.CommandBar.BuiltIn
ms.assetid: f7e4c581-2019-9fca-5e9e-15db4d656269
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.BuiltIn property (Office)

Gets **True** if the specified command bar is a built-in command bar of the container application. Returns **False** if it is a custom command bar. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BuiltIn**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Return value

Boolean


## Example

This example deletes all custom command bars that aren't visible.


```vb
foundFlag = False  
deletedBars = 0 
For Each bar In CommandBars 
    If (bar.BuiltIn = False) And (bar.Visible = False) Then 
        bar.Delete 
        foundFlag = True  
        deletedBars = deletedBars + 1 
    End If 
Next 
If Not foundFlag Then 
    MsgBox "No command bars have been deleted." 
Else 
    MsgBox deletedBars & " custom command bar(s) deleted." 
End If
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]