---
title: CommandBars.GetLabelMso method (Office)
keywords: vbaof11.chm2022
f1_keywords:
- vbaof11.chm2022
ms.prod: office
api_name:
- Office.CommandBars.GetLabelMso
ms.assetid: 1ab6f700-e3c3-a89d-790f-10c27a6b495c
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetLabelMso method (Office)

Returns the label of the control identified by the _idMso_ parameter as a **String**.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**GetLabelMso** (_idMso_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Return value

String


## Example

The following sample returns the string **Paste**.

```vb
Application.CommandBars.GetLabelMso("Paste")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]