---
title: Document.Compatibility property (Word)
keywords: vbawd10.chm158007351
f1_keywords:
- vbawd10.chm158007351
ms.prod: word
api_name:
- Word.Document.Compatibility
ms.assetid: f41979a3-8650-1807-9cf0-d1e5fdf3a49b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Compatibility property (Word)

 **True** if the compatibility option specified by the Type argument is enabled. Compatibility options affect how a document is displayed in Microsoft Word. Read/write **Boolean**.


## Syntax

_expression_. `Compatibility`( `_Type_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdCompatibility**|The compatibility option.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example enables the Suppress Space Before after a hard page or column break option on the Compatibility tab in the Options dialog box (Tools menu) for the active document.


```vb
ActiveDocument.Compatibility(wdSuppressSpBfAfterPgBrk) = True
```

This example switches the Don't add automatic tab stop for hanging indent option on or off.




```vb
ActiveDocument.Compatibility(wdNoTabHangIndent) = Not _ 
 ActiveDocument.Compatibility(wdNoTabHangIndent)
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]