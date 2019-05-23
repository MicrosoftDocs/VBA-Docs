---
title: Selection.StartIsActive property (Word)
keywords: vbawd10.chm158663060
f1_keywords:
- vbawd10.chm158663060
ms.prod: word
api_name:
- Word.Selection.StartIsActive
ms.assetid: 734e5368-dd6e-d84a-b445-30540948ac7a
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.StartIsActive property (Word)

 **True** if the beginning of the selection is active. Read/write **Boolean**.


## Syntax

_expression_. `StartIsActive`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

If the selection is not collapsed to an insertion point, either the beginning or the end of the selection is active. The active end of the selection moves when you call the following methods:  **[EndKey](Word.Selection.EndKey.md)**, **[Extend](Word.Selection.Extend.md)** (with the Characters argument), **[HomeKey](Word.Selection.HomeKey.md)**, **[MoveDown](Word.Selection.MoveDown.md)**, **[MoveLeft](Word.Selection.MoveLeft.md)**, **[MoveRight](Word.Selection.MoveRight.md)**, and **[MoveUp](Word.Selection.MoveUp.md)**.

This property is equivalent to using the  **[Flags](Word.Selection.Flags.md)** property with the **wdSelStartActive** constant. However, using the **Flags** property requires binary operations, which are more complicated than using the **StartIsActive** property.


## Example

This example extends the current selection through the next two words. To make sure that any currently selected text stays selected during the extension, the end of the selection is made active first. (For example, if the first three words of this paragraph were selected but the start of the selection were active, the  **MoveRight** method call would cancel the selection of the first two words.)


```vb
With Selection 
 .StartIsActive = False 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
End With
```

Here is the same example using the  **Flags** property. This solution is problematic because you can only deactivate a **Flags** property setting by overwriting it with an unrelated value.




```vb
With Selection 
 If (.Flags And wdSelStartActive) = wdSelStartActive Then _ 
 .Flags = wdSelReplace 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
End With
```

Here is the same example using the  **MoveEnd** method, which eliminates the need to check which end of the selection is active.




```vb
With Selection 
 .MoveEnd Unit:=wdWord, Count:=2 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]