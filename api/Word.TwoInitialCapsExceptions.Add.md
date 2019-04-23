---
title: TwoInitialCapsExceptions.Add method (Word)
keywords: vbawd10.chm155451493
f1_keywords:
- vbawd10.chm155451493
ms.prod: word
api_name:
- Word.TwoInitialCapsExceptions.Add
ms.assetid: 46aa7bea-ada5-63a8-1461-5c0a058a0981
ms.date: 06/08/2017
localization_priority: Normal
---


# TwoInitialCapsExceptions.Add method (Word)

Returns a  **TwoInitialCapsException** object that represents a new exception added to the list of AutoCorrect exceptions.


## Syntax

_expression_.**Add** (_Name_)

_expression_ Required. A variable that represents a '[TwoInitialCapsExceptions](Word.twoinitialcapsexceptions.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word with two initial capital letters that you want Microsoft Word to overlook.|

## Return value

TwoInitialCapsException


## Remarks

If the  **TwoInitialCapsAutoAdd** property is **True**, words are automatically added to the list of initial-capital exceptions.


## Example

This example adds the abbreviation addr. to the list of first-letter exceptions.


```vb
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```

This example adds MSOffice to the list of initial-capital exceptions.




```vb
AutoCorrect.TwoInitialCapsExceptions.Add Name:="MSOffice"
```

This example adds myCompany to the list of other corrections exceptions.




```vb
AutoCorrect.OtherCorrectionsExceptions.Add Name:="myCompany"
```


## See also


[TwoInitialCapsExceptions Collection Object](Word.twoinitialcapsexceptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]