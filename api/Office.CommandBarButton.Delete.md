---
title: CommandBarButton.Delete method (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Delete
ms.assetid: af94a209-b651-442f-8fa3-3a6436833d15
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Delete method (Office)

Deletes the **CommandBarButton** object from its collection.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Delete**(_Temporary_)

_expression_ Required. A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Temporary_|Optional|**Variant**|**True** to delete the control for the current session. The application will display the control again in the next session.|

## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]