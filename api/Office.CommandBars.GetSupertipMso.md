---
title: CommandBars.GetSupertipMso method (Office)
keywords: vbaof11.chm2024
f1_keywords:
- vbaof11.chm2024
ms.prod: office
api_name:
- Office.CommandBars.GetSupertipMso
ms.assetid: e116402f-bbb7-8cd3-6305-7daf85feb514
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetSupertipMso method (Office)

Returns the supertip of the control identified by the _idMso_ parameter as a **String**.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**GetSupertipMso** (_idMso_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Return value

String


## Example

The following sample returns the string "Cut the selection from the document and put it on the Clipboard."

```vb
Application.CommandBars.GetSupertipMso("Cut")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]