---
title: Document.ReadabilityStatistics property (Word)
keywords: vbawd10.chm158007392
f1_keywords:
- vbawd10.chm158007392
ms.prod: word
api_name:
- Word.Document.ReadabilityStatistics
ms.assetid: e9da9d92-bc1f-d575-07b1-3eae2749a9e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReadabilityStatistics property (Word)

Returns a  **ReadabilityStatistics** collection that represents the readability statistics for the specified document or range. Read-only.


## Syntax

_expression_. `ReadabilityStatistics`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays each readability statistic, along with its value, for document one.


```vb
For Each rs In Documents(1).ReadabilityStatistics 
 Msgbox rs.Name & " - " & rs.Value 
Next rs
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]