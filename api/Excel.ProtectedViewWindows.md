---
title: ProtectedViewWindows Object (Excel)
keywords: vbaxl10.chm912072
f1_keywords:
- vbaxl10.chm912072
ms.prod: excel
api_name:
- Excel.ProtectedViewWindows
ms.assetid: c280b1c5-c605-6453-3604-3a409a8289d0
ms.date: 06/08/2017
---


# ProtectedViewWindows Object (Excel)

A collection of the  **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** objects that represent all the **Protected View** windows that are currently open in the application.


## Remarks

Use the  **[ProtectedViewWindows](Excel.Application.ProtectedViewWindows.md)** property to return the **ProtectedViewWindows** collection.


## Example

The following code example displays the number of  **Protected View** windows that are open.


```vb
MsgBox "There are " & ProtectedViewWindows.Count & _ 
 " Protected View windows currently open."
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

