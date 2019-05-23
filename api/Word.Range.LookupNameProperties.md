---
title: Range.LookupNameProperties method (Word)
keywords: vbawd10.chm157155505
f1_keywords:
- vbawd10.chm157155505
ms.prod: word
api_name:
- Word.Range.LookupNameProperties
ms.assetid: a3a0facf-898a-d8c9-033a-b48416b53266
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.LookupNameProperties method (Word)

Looks up a name in the global address book list and displays the  **Properties** dialog box, which includes information about the specified name.


## Syntax

_expression_. `LookupNameProperties`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

If this method finds more than one match, it displays the  **Check Names** dialog box.


## Example

This example looks up the selected name in the address book and displays the  **Properties** dialog box for that person.


```vb
Selection.Range.LookupNameProperties
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]