---
title: FirstLetterExceptions.Add method (Word)
keywords: vbawd10.chm155582565
f1_keywords:
- vbawd10.chm155582565
ms.prod: word
api_name:
- Word.FirstLetterExceptions.Add
ms.assetid: 66ed8423-2c64-e924-2b34-45daea68efac
ms.date: 06/08/2017
localization_priority: Normal
---


# FirstLetterExceptions.Add method (Word)

Returns a  **FirstLetterException** object that represents a new exception added to the list of AutoCorrect exceptions.


## Syntax

_expression_.**Add** (_Name_)

_expression_ Required. A variable that represents a '[FirstLetterExceptions](Word.firstletterexceptions.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word with two initial capital letters that you want Microsoft Word to overlook.|

## Return value

FirstLetterException


## Remarks

If the **FirstLetterAutoAdd** property is **True**, abbreviations are automatically added to the list of first-letter exceptions.


## Example

This example adds the abbreviation addr. to the list of first-letter exceptions.


```vb
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```


## See also


[FirstLetterExceptions Collection Object](Word.firstletterexceptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]