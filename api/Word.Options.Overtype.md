---
title: Options.Overtype property (Word)
keywords: vbawd10.chm162988098
f1_keywords:
- vbawd10.chm162988098
ms.prod: word
api_name:
- Word.Options.Overtype
ms.assetid: 2538fee5-3571-3fae-06d0-f6c3533bb121
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.Overtype property (Word)

 **True** if Overtype mode is active. Read/write **Boolean**.


## Syntax

_expression_. `Overtype`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

In Overtype mode, the characters you type replace existing characters one by one. When Overtype isn't active, the characters you type move existing text to the right.


## Example

If Overtype mode is active, this example displays a message box asking whether Overtype should be deactivated. If the user clicks the Yes button, Overtype mode is made inactive.


```vb
If Options.Overtype = True Then 
 aButton = MsgBox("Overtype is on. Turn off?", 4) 
 If aButton = vbYes Then Options.Overtype = False 
End If
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]