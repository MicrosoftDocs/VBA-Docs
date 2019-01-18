---
title: CommandBarControl.Enabled property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Enabled
ms.assetid: 74105bf5-96a0-09ea-bb00-ef102705372c
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.Enabled property (Office)

Gets or sets a **Boolean** value specifying if the **CommandBarControl** is enabled. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Remarks

For command bars, setting this property to **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]