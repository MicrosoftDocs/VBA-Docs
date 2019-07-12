---
title: Document.RenameCurrentScope method (Visio)
keywords: vis_sdr.chm10550185
f1_keywords:
- vis_sdr.chm10550185
ms.prod: visio
api_name:
- Visio.Document.RenameCurrentScope
ms.assetid: 08aff947-e876-29b8-e910-e2a3b42e5d0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.RenameCurrentScope method (Visio)

Renames the top-level open undo scope.


## Syntax

_expression_.**RenameCurrentScope** (_bstrScopeName_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrScopeName_|Required| **String**|The new name of the undo scope.|

## Return value

Nothing


## Remarks

The new name assigned to the undo scope appears on the **Undo** menu as the item name. If there is no open undo scope, the **RenameCurrentScope** method raises an exception.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]