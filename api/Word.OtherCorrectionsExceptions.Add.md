---
title: OtherCorrectionsExceptions.Add method (Word)
keywords: vbawd10.chm165609573
f1_keywords:
- vbawd10.chm165609573
ms.prod: word
api_name:
- Word.OtherCorrectionsExceptions.Add
ms.assetid: 0bdb30c5-72f0-3dae-e0c5-b2ea48157626
ms.date: 06/08/2017
localization_priority: Normal
---


# OtherCorrectionsExceptions.Add method (Word)

Returns an  **OtherCorrectionsException** object that represents a new exception added to the list of AutoCorrect exceptions.


## Syntax

_expression_.**Add** (_Name_)

_expression_ Required. A variable that represents an '[OtherCorrectionsExceptions](Word.othercorrectionsexceptions.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word that you want Word to overlook.|

## Return value

OtherCorrectionsException


## Remarks

If the  **OtherCorrectionsAutoAdd** property is **True**, words are automatically added to the list of other corrections exceptions.


## Example

This example adds myCompany to the list of other corrections exceptions.


```vb
AutoCorrect.OtherCorrectionsExceptions.Add Name:="myCompany"
```


## See also


[OtherCorrectionsExceptions Collection Object](Word.othercorrectionsexceptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]