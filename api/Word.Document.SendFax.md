---
title: Document.SendFax method (Word)
keywords: vbawd10.chm158007552
f1_keywords:
- vbawd10.chm158007552
ms.prod: word
api_name:
- Word.Document.SendFax
ms.assetid: d7a1636b-1fc2-9bfd-e7f6-828a745c06d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SendFax method (Word)

Sends the specified document as a fax, without any user interaction.


## Syntax

_expression_. `SendFax`( `_Address_` , `_Subject_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Address_|Required| **String**|The recipient's fax number.|
| _Subject_|Optional| **Variant**|The text for the subject line. The character limit is 255.|

## Example

This example sends the active document as a fax.


```vb
ActiveDocument.SendFax Address:="12065551234", _ 
 Subject:="Important Fax"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]