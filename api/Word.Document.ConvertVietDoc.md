---
title: Document.ConvertVietDoc method (Word)
keywords: vbawd10.chm158007743
f1_keywords:
- vbawd10.chm158007743
ms.prod: word
api_name:
- Word.Document.ConvertVietDoc
ms.assetid: d03f0ad4-0e40-45a7-5189-1cbfa7328b2c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ConvertVietDoc method (Word)

Reconverts a Vietnamese document to Unicode using a code page other than the default.


## Syntax

_expression_. `ConvertVietDoc`( `_CodePageOrigin_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CodePageOrigin_|Required| **Long**|The original code page used to encode the document.|

## Remarks

Use the  **ConvertVietDoc** method if you want a document to be viewable on another computer or platform.


## Example

This example converts the active document from the Vietnamese ABC code page to Unicode. This example assumes that the active document is encoded using the Vietnamese ABC code page.


```vb
Sub ConvertToVietCodePage() 
 ActiveDocument.ConvertVietDoc CodePageOrigin:=5 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]