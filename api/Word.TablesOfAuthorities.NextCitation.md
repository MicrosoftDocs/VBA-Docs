---
title: TablesOfAuthorities.NextCitation method (Word)
keywords: vbawd10.chm152174695
f1_keywords:
- vbawd10.chm152174695
ms.prod: word
api_name:
- Word.TablesOfAuthorities.NextCitation
ms.assetid: c0bfde51-ce49-1570-9599-515b43875dec
ms.date: 06/08/2017
localization_priority: Normal
---


# TablesOfAuthorities.NextCitation method (Word)

Finds and selects the next instance of the text specified by the ShortCitation parameter.


## Syntax

_expression_. `NextCitation`( `_ShortCitation_` )

_expression_ Required. A variable that represents a '[TablesOfAuthorities](Word.tablesofauthorities.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShortCitation_|Required| **String**|The text of the short citation.|

## Example

This example selects the next citation in the active document that begins with "in re".


```vb
ActiveDocument.TablesOfAuthorities.NextCitation _ 
 ShortCitation:="in re"
```


## See also


[TablesOfAuthorities Collection Object](Word.tablesofauthorities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]