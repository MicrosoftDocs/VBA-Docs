---
title: Document.TransformDocument method (Word)
keywords: vbawd10.chm158007796
f1_keywords:
- vbawd10.chm158007796
ms.prod: word
api_name:
- Word.Document.TransformDocument
ms.assetid: 5829a16f-b514-479f-c227-359123611970
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.TransformDocument method (Word)

Applies the specified Extensible Stylesheet Language Transformation (XSLT) file to the specified document and replaces the document with the results.


## Syntax

_expression_. `TransformDocument`( `_Path_` , `_DataOnly_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path for the XSLT to use.|
| _DataOnly_|Optional| **Boolean**| **True** applies the transformation only to the data in the document, excluding Microsoft Word XML. **False** applies the transform to the entire document, including Word XML. Default value is **True**.|

## Example

The following example transforms the active document using the specified XSLT file.


```vb
ActiveDocument.TransformDocument _ 
 ("c:\schemas\simplesample.xslt")
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]