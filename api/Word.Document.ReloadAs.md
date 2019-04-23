---
title: Document.ReloadAs method (Word)
keywords: vbawd10.chm158007627
f1_keywords:
- vbawd10.chm158007627
ms.prod: word
api_name:
- Word.Document.ReloadAs
ms.assetid: 52cab019-3084-e488-8727-24c5fd4650ce
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReloadAs method (Word)

Reloads a document based on an HTML document, using the specified document encoding.


## Syntax

_expression_. `ReloadAs`( `_Encoding_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Encoding_|Required| **MsoEncoding**|Specifies the encoding to use when reloading the document.|

## Example

This example reloads the current document with Cyrillic encoding.


```vb
ActiveDocument.ReloadAs msoEncodingCyrillic
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]