---
title: Global.Repeat method (Word)
keywords: vbawd10.chm163119409
f1_keywords:
- vbawd10.chm163119409
ms.prod: word
api_name:
- Word.Global.Repeat
ms.assetid: 23e2e300-cc01-cd9d-f761-0113a07267bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Repeat method (Word)

Repeats the most recent editing action one or more times. Returns  **True** if the commands were repeated successfully.


## Syntax

_expression_.**Repeat** (_Times_)

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Times_|Optional| **Variant**|The number of times you want to repeat the last command.|

## Return value

Boolean


## Remarks

Using this method is the equivalent to using the **Repeat** command on the **Edit** menu.


## Example

This example inserts the text "Hello" followed by two paragraphs (the second typing action is repeated once).


```vb
Selection.TypeText "Hello" 
Selection.TypeParagraph 
Repeat
```

This example repeats the last command three times (if it can be repeated).




```vb
On Error Resume Next 
If Repeat(3) = True Then StatusBar = "Action repeated"
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]