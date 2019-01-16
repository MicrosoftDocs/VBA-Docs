---
title: ProtectedViewWindow object (Word)
keywords: vbawd10.chm3536
f1_keywords:
- vbawd10.chm3536
ms.prod: word
api_name:
- Word.ProtectedViewWindow
ms.assetid: d77e80e7-c54e-5954-1586-dacd3c9f7434
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow object (Word)

Represents a protected view window.


## Remarks

Documents displayed in a protected view window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and Data Connections.

Use [ProtectedViewWindows](Word.ProtectedViewWindows.md)( _Index_ ), where _Index_ is the index number to return a single **ProtectedViewWindow** object.


## Example

The index number represents the position of the protected view window in the  **ProtectedViewWindows** collection.. The following code example returns the first protected view window.


```vb
Dim pvWindow As ProtectedViewWindow 
 
Set pvWindow = ProtectedViewWindows(1) 

```


## See also


[Word Object Model Reference](./overview/Word/object-model.md)


