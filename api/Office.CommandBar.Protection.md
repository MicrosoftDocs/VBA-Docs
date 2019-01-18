---
title: CommandBar.Protection property (Office)
keywords: vbaof11.chm3015
f1_keywords:
- vbaof11.chm3015
ms.prod: office
api_name:
- Office.CommandBar.Protection
ms.assetid: 59f9e9d3-251c-93a6-fa49-75fa7c4f6659
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Protection property (Office)

Gets or sets an **[msoBarProtection](office.msobarprotection.md)** constant representing the way a command bar is protected from user customization. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Protection**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

Using the constant **msoBarNoCustomize** prevents users from accessing the **Add or Remove Buttons** menu (this menu enables users to customize a toolbar).


## Example

This example steps through the collection of command bars to find the command bar named **Forms**. If this command bar is found, its docking state is protected and it is made visible.


```vb
foundFlag =  False 
For i = 1 To CommandBars.Count 
    If CommandBars(i).Name = "Forms" Then 
            CommandBars(i).Protection = msoBarNoChangeDock 
            CommandBars(i).Visible = True  
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