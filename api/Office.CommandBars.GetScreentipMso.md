---
title: CommandBars.GetScreentipMso method (Office)
keywords: vbaof11.chm2023
f1_keywords:
- vbaof11.chm2023
ms.prod: office
api_name:
- Office.CommandBars.GetScreentipMso
ms.assetid: 23411622-2b35-0c0e-9373-9bc75c5e433e
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetScreentipMso method (Office)

Returns the screentip of the control identified by the _idMso_ parameter as a **String**.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**GetScreentipMso** (_idMso_)

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
Application.CommandBars.GetScreentipMso("Paste")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]