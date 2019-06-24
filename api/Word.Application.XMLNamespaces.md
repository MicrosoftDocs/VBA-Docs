---
title: Application.XMLNamespaces property (Word)
keywords: vbawd10.chm158335439
f1_keywords:
- vbawd10.chm158335439
ms.prod: word
api_name:
- Word.Application.XMLNamespaces
ms.assetid: e7eac332-f805-5ceb-682c-482565ff0786
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.XMLNamespaces property (Word)

Returns an  **** collection that represents the XML schemas in the Schema Library.


## Syntax

_expression_. `XMLNamespaces`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

The following example returns the first schema in the Schema Library.


```vb
Dim objSchema As XMLNamespace 
 
Set objSchema = Application.XMLNamespaces.Item(1)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]