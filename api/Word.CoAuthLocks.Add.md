---
title: CoAuthLocks.Add method (Word)
keywords: vbawd10.chm180486146
f1_keywords:
- vbawd10.chm180486146
ms.prod: word
api_name:
- Word.CoAuthLocks.Add
ms.assetid: e66aed3e-b097-31c5-3b2a-748e278c3b61
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLocks.Add method (Word)

Returns a  **[CoAuthLock](Word.CoAuthLock.md)** object that represents a lock added to a specified range.


## Syntax

_expression_.**Add** (_Range_, _Type_)

 _expression_ An expression that returns a '[CoAuthLocks](Word.CoAuthLocks.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Variant**|Specifies the document range locked by the  **[CoAuthLock](Word.CoAuthLock.md)** object. This parameter may be a **Paragraph**, **Column**, **Cell**, **Row**, **Table**, **Range**, or **Selection** object.|
| _Type_|Optional| **[WdLockType](Word.WdLockType.md)**|Specifies the type of lock. The  **WdLockType** specified can only be **wdLockEphemeral** or **WdLockReservation**|

## Return value

 **CoAuthLock**


## Remarks

The following code example adds a reservation lock to the first paragraph in the active document.


> [!NOTE] 
> By default, if no arguments are given in the call to the  **CoAuthLocks.Add** method, a reservation lock is placed on the paragraph that contains the insertion point.


## Example


```vb
ActiveDocument.CoAuthoring.Locks.Add(ActiveDocument.Paragraphs(1).Range, wdLockReservation)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]