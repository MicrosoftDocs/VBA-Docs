---
title: Application.Repeat method (Word)
keywords: vbawd10.chm158335281
f1_keywords:
- vbawd10.chm158335281
ms.prod: word
api_name:
- Word.Application.Repeat
ms.assetid: 811e9f1c-cbdc-01dc-1e76-5521976943ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Repeat method (Word)

Repeats the most recent editing action one or more times. Returns  **True** if the commands were repeated successfully.


## Syntax

_expression_.**Repeat** (_Times_)

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]