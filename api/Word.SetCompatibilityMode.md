---
title: Document.SetCompatibilityMode method (Word)
keywords: vbawd10.chm158007867
f1_keywords:
- vbawd10.chm158007867
ms.prod: word
api_name:
- Word.SetCompatibilityMode
ms.assetid: f167a640-340e-56ed-34c0-0c3dbff8575a
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SetCompatibilityMode method (Word)

Sets the compatibility mode for the document.


## Syntax

_expression_. `SetCompatibilityMode`( `_Mode_` )

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Mode_|Required| **Long**|Specifies which version of Word to approximate. Use a constant from the [WdCompatibilityMode](Word.WdCompatibilityMode.md) enumeration as an argument for this parameter.|

## Remarks

When you open a document in Word that was created in a previous version of Word, Compatibility Mode is turned on. Compatibility Mode ensures that no new or enhanced features in Word are available while working with a document, so that people who edit the document using previous versions of Word will have full editing capabilities.


## Example

The following code example puts Word in Word 2003 Compatibility Mode.


```vb
ActiveDocument.SetCompatibilityMode (wdWord2003)
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]