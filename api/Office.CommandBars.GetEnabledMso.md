---
title: CommandBars.GetEnabledMso method (Office)
keywords: vbaof11.chm2019
f1_keywords:
- vbaof11.chm2019
ms.prod: office
api_name:
- Office.CommandBars.GetEnabledMso
ms.assetid: 68af6404-53ee-4c69-51fa-4d489736d228
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetEnabledMso method (Office)

Returns **True** if the control identified by the _idMso_ parameter is enabled.


## Syntax

_expression_.**GetEnabledMso** (_idMso_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Return value

Boolean

## Example

The following sample returns **True** if the **Bold** button is enabled.

```vb
Application.CommandBars.GetEnabledMso("Bold")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]