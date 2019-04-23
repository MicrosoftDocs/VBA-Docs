---
title: XMLNamespace.URI property (Word)
keywords: vbawd10.chm2293762
f1_keywords:
- vbawd10.chm2293762
ms.prod: word
api_name:
- Word.URI
ms.assetid: 540997ad-ead3-dcda-c5c7-ddfc7877fedc
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNamespace.URI property (Word)

Returns a  **String** that represents the Uniform Resource Identifier (URI) of the associated namespace.


## Syntax

_expression_. `URI`

 _expression_ An expression that returns an '[XMLNamespace](Word.XMLNamespace.md)' object.


## Example

The following example displays the URI for the first schema in the Schema Library.


```vb
MsgBox Application.XMLNamespaces(1).URI
```


## See also


[XMLNamespace Object](Word.XMLNamespace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]