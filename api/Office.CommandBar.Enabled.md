---
title: CommandBar.Enabled property (Office)
keywords: vbaof11.chm3005
f1_keywords:
- vbaof11.chm3005
ms.prod: office
api_name:
- Office.CommandBar.Enabled
ms.assetid: 4a332d30-4aa9-1355-2d26-0d4f0529d488
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Enabled property (Office)

Gets or sets a **Boolean** value that specifies whether the specified **CommandBar** is enabled. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

For command bars, setting this property to **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]