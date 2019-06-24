---
title: Application.CheckGrammar method (Word)
keywords: vbawd10.chm158335299
f1_keywords:
- vbawd10.chm158335299
ms.prod: word
api_name:
- Word.Application.CheckGrammar
ms.assetid: 4675bda9-c31d-efdc-4def-38bfdeb200e4
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CheckGrammar method (Word)

Checks a string for grammatical errors. Returns a  **Boolean** to indicate whether the string contains grammatical errors. **True** if the string contains no errors.


## Syntax

_expression_. `CheckGrammar`( `_String_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _String_|Required| **String**|The string you want to check for grammatical errors.|

## Return value

Boolean


## Example

This example displays the result of a grammar check on the selection.


```vb
strPass = Application.CheckGrammar(String:=Selection.Text) 
MsgBox "Selection is grammatically correct: " & strPass
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]