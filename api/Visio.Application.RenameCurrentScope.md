---
title: Application.RenameCurrentScope method (Visio)
keywords: vis_sdr.chm10050815
f1_keywords:
- vis_sdr.chm10050815
ms.prod: visio
api_name:
- Visio.Application.RenameCurrentScope
ms.assetid: 0ccd9c6b-704c-b956-5ea9-4f1ed01baee7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RenameCurrentScope method (Visio)

Renames the top-level open undo scope.


## Syntax

_expression_.**RenameCurrentScope** (_bstrScopeName_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrScopeName_|Required| **String**|The new name of the undo scope.|

## Return value

Nothing


## Remarks

The new name assigned to the undo scope appears on the **Undo** menu as the item name. If there is no open undo scope, the **RenameCurrentScope** method raises an exception.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]