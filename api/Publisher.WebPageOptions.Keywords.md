---
title: WebPageOptions.Keywords property (Publisher)
keywords: vbapb10.chm544772
f1_keywords:
- vbapb10.chm544772
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.Keywords
ms.assetid: 8dd7b073-747e-a6f6-a20d-0b3e3d9a27b8
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.Keywords property (Publisher)

Returns or sets a **String** that represents the keywords for a webpage within a web publication. Read/write.


## Syntax

_expression_.**Keywords**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Return value

String


## Example

The following example sets the keywords for page four of the active publication.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .Keywords = "software, hardware, computers" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]