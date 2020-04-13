---
title: Source.Field property (Word)
keywords: vbawd10.chm140836968
f1_keywords:
- vbawd10.chm140836968
ms.prod: word
api_name:
- Word.Source.Field
ms.assetid: fd6689d4-a042-4ca2-fddd-d048fe8c3a93
ms.date: 06/08/2017
localization_priority: Normal
---


# Source.Field property (Word)

Returns a  **String** that represents the value of a field in a bibliography source. Read-only.


## Syntax

_expression_. `Field`( `_Name_` )

 _expression_ An expression that returns a [Source](./Word.Source.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the field for which to retrieve the value.|

## Remarks

The name of the field corresponds to the name of the corresponding XML element in the resulting XML for a bibliography source. You can use the **[XML](Word.Source.XML.md)** property to return the XML for a bibliography source. For more information, see [Working with Bibliographies](../word/Concepts/Working-with-Word/working-with-bibliographies.md).


## See also


[Source Object](Word.Source.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]