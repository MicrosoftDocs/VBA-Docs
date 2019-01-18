---
title: Font.Shadow property (Word)
keywords: vbawd10.chm156369042
f1_keywords:
- vbawd10.chm156369042
ms.prod: word
api_name:
- Word.Font.Shadow
ms.assetid: e81f8b86-7f60-7852-6c72-7b01de832447
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.Shadow property (Word)

 **True** if the specified font is formatted as shadowed. Read/write **Long**.


## Syntax

 _expression_. `Shadow`

 _expression_ Required. A variable that represents a '[Font](Word.Font.md)' object.


## Remarks

This property can be  **True** , **False** , or **wdUndefined**.


## Example

This example applies shadow and bold formatting to the selection.


```vb
If Selection.Type = wdSelectionNormal Then 
 With Selection.Font 
 .Shadow = True 
 .Bold = True 
 End With 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]