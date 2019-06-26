---
title: Selection.ReplaceShape method (Visio)
ms.prod: visio
ms.assetid: dc278901-77ce-e1fe-c44f-f464bbb1c360
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ReplaceShape method (Visio)

Replaces the specified selection with one or more instances of the master passed as the first parameter, and returns an array containing the new shape or shapes.


## Syntax

_expression_.**ReplaceShape** (_MasterOrMasterShortcutToDrop_, _ReplaceFlags_)

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MasterOrMasterShortcutToDrop_|Required|UNKNOWN|Specifies the replacement shape or shapes to drop. Must be either a **[Master](Visio.Master.md)** or **[MasterShortcut](Visio.MasterShortcut.md)** object.|
| _ReplaceFlags_|Optional|INT32|Specifies the properties of the original shape or shapes to retain in the new shape or shapes. Possible values include any of the **[VisReplaceFlags](Visio.visreplaceflags.md)** constants, and certain combinations of those constants. See Remarks for more information.|

## Return value

**SAFE-ARRAY**


## Remarks

Allowable values to pass for the _ReplaceFlags_ parameter include either **visReplaceShapeDefault** or any combination of one or more of the remaining four flags.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]