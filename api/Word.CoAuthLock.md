---
title: CoAuthLock object (Word)
keywords: vbawd10.chm3968
f1_keywords:
- vbawd10.chm3968
ms.prod: word
api_name:
- Word.CoAuthLock
ms.assetid: 3efa12b0-1079-c6df-20c1-a66398161c8e
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthLock object (Word)

Represents a lock within the document. The **CoAuthLock** object is a member of the **[CoAuthLocks](Word.CoAuthLocks.md)** collection.


## Remarks

Use  **Locks** (_index_), where _index_ is the index number, to return a **CoAuthLock** object. When adding a **CoAuthLock** object, use the **[WdLockType](Word.WdLockType.md)** enumeration to specify the type of lock.


## Example

The following code example returns the first lock in the active document.


```vb
Dim myLock as CoAuthLock 
 
Set myLock = ActiveDocument.CoAuthoring.Locks(1)
```

The following code example adds a reservation lock on the third paragraph in the active document. Reservation locks are explicitly created by a document author and are persisted across explicit save actions (locks of type  **wdLockEphemeral** do not persist across explicit saves). You can add locks with a with a lock type of **wdLockReservation** using the Word ribbon. For example, you can create a reservation lock on a selected paragraph range using **Block Authors** on the **Review** tab.




```vb
ActiveDocument.CoAuthoring.Locks.Add(ActiveDocument.Paragraphs(3), wdLockReservation)
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]