---
title: Document.PrintPageBackgrounds property (Publisher)
keywords: vbapb10.chm196743
f1_keywords:
- vbapb10.chm196743
ms.prod: publisher
api_name:
- Publisher.Document.PrintPageBackgrounds
ms.assetid: 6d1d6e6a-fd66-2afa-2172-4a6552d5cce4
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.PrintPageBackgrounds property (Publisher)

Returns or sets **True** to include page backgrounds when printing pages from the specified publication. Default is **True**. Read/write **Boolean**.


## Syntax

_expression_.**PrintPageBackgrounds**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Boolean


## Remarks

Use the **[PageBackground](Publisher.PageBackground.md)** object to create, alter, or delete the background of a specified page.


## Example

The following example sets page backgrounds to print for the active publication.

```vb
ActiveDocument.PrintPageBackgrounds = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]