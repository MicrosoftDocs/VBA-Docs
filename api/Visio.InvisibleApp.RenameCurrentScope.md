---
title: InvisibleApp.RenameCurrentScope method (Visio)
keywords: vis_sdr.chm17550815
f1_keywords:
- vis_sdr.chm17550815
ms.prod: visio
api_name:
- Visio.InvisibleApp.RenameCurrentScope
ms.assetid: f057117c-5565-60a8-2c19-d30f6c6b5c28
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.RenameCurrentScope method (Visio)

Renames the top-level open undo scope.


## Syntax

_expression_.**RenameCurrentScope** (_bstrScopeName_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrScopeName_|Required| **String**|The new name of the undo scope.|

## Return value

**Nothing**


## Remarks

The new name assigned to the undo scope appears on the  **Undo** menu as the item name. If there is no open undo scope, the **RenameCurrentScope** method raises an exception.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]