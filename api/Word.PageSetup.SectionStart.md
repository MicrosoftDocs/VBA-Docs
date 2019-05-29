---
title: PageSetup.SectionStart property (Word)
keywords: vbawd10.chm158400626
f1_keywords:
- vbawd10.chm158400626
ms.prod: word
api_name:
- Word.PageSetup.SectionStart
ms.assetid: 2fa3d589-82e7-9f9a-7b0e-f8761dd60a9a
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.SectionStart property (Word)

Returns or sets the type of section break for the specified object. Read/write  **WdSectionStart**.


## Syntax

_expression_. `SectionStart`

_expression_ Required. A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example changes the type of section break to continuous for all sections in the active document.


```vb
ActiveDocument.PageSetup.SectionStart = wdSectionContinuous
```

This example returns the type of section break used at the beginning of the second section in MyDoc.doc and applies it to all the sections in the active document.




```vb
mytype = Documents("MyDoc.doc").Sections(2).PageSetup.SectionStart 
ActiveDocument.PageSetup.SectionStart = mytype
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]