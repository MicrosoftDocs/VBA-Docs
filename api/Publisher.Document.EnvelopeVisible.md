---
title: Document.EnvelopeVisible Property (Publisher)
keywords: vbapb10.chm196618
f1_keywords:
- vbapb10.chm196618
ms.prod: publisher
api_name:
- Publisher.Document.EnvelopeVisible
ms.assetid: 65423c1f-e61b-3c83-4bff-ddd278d97238
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.EnvelopeVisible Property (Publisher)

Returns or sets a  **Boolean** indicating whether the email message header is visible in the publication window. Read/write.


## Syntax

 _expression_. **EnvelopeVisible**

 _expression_ A variable that represents an  **Document** object.


## Return value

Boolean


## Example

This example displays the email message header for the active publication.


```vb
ActiveDocument.EnvelopeVisible = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]