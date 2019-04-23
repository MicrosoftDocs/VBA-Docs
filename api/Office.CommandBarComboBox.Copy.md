---
title: CommandBarComboBox.Copy method (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Copy
ms.assetid: 15eb757c-bb07-cd98-ff9e-1810db4f475c
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.Copy method (Office)

Copies a command bar combo box control to an existing command bar.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Copy** (_Bar_, _Before_)

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Bar_|Optional|**Variant**| A **CommandBar** object that represents the destination command bar. If this argument is omitted, the control is copied to the command bar where the control already exists.|
| _Before_|Optional|**Variant**|A number that indicates the position for the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is copied to the end of the command bar.|

## Return value

CommandBarControl


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]