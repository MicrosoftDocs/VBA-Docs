---
title: CommandBars.ExecuteMso method (Office)
keywords: vbaof11.chm2018
f1_keywords:
- vbaof11.chm2018
ms.prod: office
api_name:
- Office.CommandBars.ExecuteMso
ms.assetid: 6f608475-7a79-48c7-abff-86d9ab07fe80
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.ExecuteMso method (Office)

Executes the control identified by the _idMso_ parameter.


## Syntax

_expression_.**ExecuteMso** (_idMso_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Remarks

This method is useful in cases where there is no object model for a particular command. Works on controls that are built-in buttons, toggleButtons, and splitButtons. On failure it returns E_InvalidArg for an invalid _idMso_, and E_Fail for controls that are not enabled or not visible.


## Example

The following sample executes the **Copy** button.

```vb
Application.CommandBars.ExecuteMso("Copy")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
