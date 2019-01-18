---
title: System.MathCoprocessorInstalled property (Word)
keywords: vbawd10.chm154468363
f1_keywords:
- vbawd10.chm154468363
ms.prod: word
api_name:
- Word.System.MathCoprocessorInstalled
ms.assetid: 77f7da63-b940-ac22-125e-596a1518b6b8
ms.date: 06/08/2017
localization_priority: Normal
---


# System.MathCoprocessorInstalled property (Word)

 **True** if a math coprocessor is installed on the system. Read-only **Boolean**.


## Syntax

 _expression_. `MathCoprocessorInstalled`

 _expression_ An expression that returns a '[System](Word.System.md)' object.


## Example

This example displays a message if a math coprocessor is installed on the system.


```vb
If System.MathCoprocessorInstalled = True Then 
 MsgBox "A math coprocessor is installed." 
End If
```


## See also


[System Object](Word.System.md)

