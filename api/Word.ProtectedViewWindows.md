---
title: ProtectedViewWindows Object (Word)
ms.prod: word
api_name:
- Word.ProtectedViewWindows
ms.assetid: 62c2f4d5-1080-548e-730b-388308144dfe
ms.date: 06/08/2017
---


# ProtectedViewWindows Object (Word)

A collection of all the [ProtectedViewWindow](Word.ProtectedViewWindow.md) objects that are currently open in Word.


## Remarks

Use the  **ProtectedViewWindows** property to return the **ProtectedViewWindows** collection.


## Example

The following code example displays the number of protected view windows that are open.


```vb
<<<<<<< HEAD
MsgBox "There are " &; ProtectedViewWindows.Count &; _ 
=======
MsgBox "There are " & ProtectedViewWindows.Count & _ 
>>>>>>> master
 " protected view windows currently open."
```


## See also


[Word Object Model Reference](./overview/Word/object-model.md)


