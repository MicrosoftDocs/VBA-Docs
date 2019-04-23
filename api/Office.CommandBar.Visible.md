---
title: CommandBar.Visible property (Office)
keywords: vbaof11.chm3020
f1_keywords:
- vbaof11.chm3020
ms.prod: office
api_name:
- Office.CommandBar.Visible
ms.assetid: c7057c83-ea8d-c167-a650-d784d5e6dd1f
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Visible property (Office)

Gets or sets the **Visible** property of the command bar. **True** if the command bar is visible. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Return value

Boolean


## Remarks

The **Visible** property for newly created custom command bars is **False** by default.

The **Enabled** property for a command bar must be set to **True** before the **Visible** property is set to **True**.


## Example

This example steps through the collection of command bars to find the **Forms** command bar. If the **Forms** command bar is found, the example makes it visible and protects its docking state.


```vb
foundFlag = False  
For Each cmdbar In CommandBars 
    If cmdbar.Name = "Forms" Then 
        cmdbar.Protection = msoBarNoChangeDock 
        cmdbar.Visible = True  
        foundFlag = True  
    End If 
Next 
If Not foundFlag Then 
    MsgBox "'Forms' command bar is not in the collection." 
End If
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]