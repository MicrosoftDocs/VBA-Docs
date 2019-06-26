---
title: Shape.ReplaceShape method (Visio)
ms.prod: visio
ms.assetid: b330a63d-4e3f-0c4d-c38c-6ee806670225
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ReplaceShape method (Visio)

Replaces the specified shape with an instance of the master passed as the first parameter, and returns the new shape.


## Syntax

_expression_.**ReplaceShape** (_MasterOrMasterShortcutToDrop_, _ReplaceFlags_)

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MasterOrMasterShortcutToDrop_|Required|UNKNOWN|Specifies the replacement shape to drop. Must be either a **[Master](Visio.Master.md)** or **[MasterShortcut](Visio.MasterShortcut.md)** object.|
| _ReplaceFlags_|Optional|INT32|Specifies the properties of the original shape to retain in the new shape. Possible values include any of the **[VisReplaceFlags](Visio.visreplaceflags.md)** constants, and certain combinations of those constants. See Remarks for more information.|

## Return value

**SHAPE**


## Remarks

Allowable values to pass for the _ReplaceFlags_ parameter include either **visReplaceShapeDefault** or any combination of one or more of the remaining four flags.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]