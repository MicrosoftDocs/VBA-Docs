---
title: XMLNode.XML property (Word)
keywords: vbawd10.chm37748741
f1_keywords:
- vbawd10.chm37748741
ms.prod: word
api_name:
- Word.XML
ms.assetid: a72c7c13-7e2f-c903-9b02-4e9af3f7db26
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.XML property (Word)

Returns a  **String** that represents the text, with or without XML markup, that is contained within an XML node. Read-only.


## Syntax

_expression_.**XML** (_DataOnly_)

_expression_ An expression that returns an [XMLNode](./Word.XMLNode.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataOnly_|Optional| **Boolean**|Specifies whether to return the XML with or without markup.  **True** returns the text contained within the XML node without XML markup. **False** returns the text with XML markup.|

## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]