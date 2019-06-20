---
title: CommandBars.GetVisibleMso method (Office)
keywords: vbaof11.chm2020
f1_keywords:
- vbaof11.chm2020
ms.prod: office
api_name:
- Office.CommandBars.GetVisibleMso
ms.assetid: ab916050-e1af-0752-9734-23d0fe27542f
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetVisibleMso method (Office)

Returns **True** if the control identified by the _idMso_ parameter is visible.

> [!NOTE] 
>  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**GetVisibleMso** (_idMso_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Return value

Boolean


## Example

The following sample returns **True** if the **Bold** button is visible.

```vb
Application.CommandBars.GetVisibleMso("Bold")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]