---
title: CommandBarControl.Delete method (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Delete
ms.assetid: eca4abea-092b-0c11-1040-7132318b1bea
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.Delete method (Office)

Deletes the **CommandBarControl** object from its collection.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Delete** (_Temporary_)

_expression_ Required. A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Temporary_|Optional|**Variant**|**True** to delete the control for the current session. The application will display the control again in the next session.|

## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]