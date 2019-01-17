---
title: Document.XMLUseXSLTWhenSaving property (Word)
keywords: vbawd10.chm158007770
f1_keywords:
- vbawd10.chm158007770
ms.prod: word
api_name:
- Word.Document.XMLUseXSLTWhenSaving
ms.assetid: b2161a4f-9169-6927-8f37-2bc7f5a0b319
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.XMLUseXSLTWhenSaving property (Word)

Returns a  **Boolean** that represents whether to save a document through an Extensible Stylesheet Language Transformation (XSLT). **True** saves a document through an XSLT.


## Syntax

 _expression_. `XMLUseXSLTWhenSaving`

 _expression_ An expression that returns a '[Document](Word.Document.md)' object.


## Remarks

When setting the XMLUseXSLTWhenSaving property to  **True** , use the **[XMLSaveThroughXSLT](Word.Document.XMLSaveThroughXSLT.md)** property to specify the path and file name of the XSLT to use.


## Example

The following example specifies that Microsoft Word will use an XSLT when saving the active document, and then specifies which XSLT to use.


```vb
ActiveDocument.XMLUseXSLTWhenSaving = True 
ActiveDocument.XMLSaveThroughXSLT = "c:\schemas\book.xslt"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]