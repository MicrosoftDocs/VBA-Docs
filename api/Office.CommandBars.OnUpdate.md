---
title: CommandBars.OnUpdate event (Office)
keywords: vbaof11.chm228001
f1_keywords:
- vbaof11.chm228001
ms.prod: office
api_name:
- Office.CommandBars.OnUpdate
ms.assetid: 4da9354b-92ed-d85e-f667-c01dfec07689
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.OnUpdate event (Office)

Occurs when any change is made to a command bar.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OnUpdate**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Remarks

The **OnUpdate** event is recognized by the **CommandBar** object and all command bar controls. The event is triggered by any change to a command bar or command bar control or any change to the state of a command bar or command bar control. These changes can occur due to a text or cell selection, for example. 

Because a large number of **OnUpdate** events can occur during normal usage, developers should exercise caution when using this event. It is strongly recommended that this event be used primarily for checking that a custom command bar has been added or removed by a **COMAddIn**.


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]