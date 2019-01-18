---
title: CommandBars.DisableAskAQuestionDropdown property (Office)
keywords: vbaof11.chm2017
f1_keywords:
- vbaof11.chm2017
ms.prod: office
api_name:
- Office.CommandBars.DisableAskAQuestionDropdown
ms.assetid: a0954aa4-256c-4a14-6bab-959a00e9367d
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.DisableAskAQuestionDropdown property (Office)

Is **True** if the **Answer Wizard** dropdown menu is enabled. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**DisableAskAQuestionDropdown**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Example

The following example switches the **DisableAskAQuestionDropdown** property on or off.


```vb
Sub ToggleQuestionDropdown() 
    With Application.CommandBars 
        If .DisableAskAQuestionDropdown =  True Then 
            .DisableAskAQuestionDropdown = False  
        Else 
            .DisableAskAQuestionDropdown = True  
        End If 
    End With 
End Sub
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]