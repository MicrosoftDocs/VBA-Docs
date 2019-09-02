---
title: Documents.Close method (Word)
keywords: vbawd10.chm158073937
f1_keywords:
- vbawd10.chm158073937
ms.prod: word
api_name:
- Word.Documents.Close
ms.assetid: 0284daf3-311e-97c9-ffc6-74d63b85fdcd
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.Close method (Word)

Closes the specified documents.


## Syntax

_expression_.**Close** (_SaveChanges_, _OriginalFormat_, _RouteDocument_)

_expression_ Required. A variable that represents a **[Documents](Word.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|Specifies the save action for the document. Can be one of the following **[WdSaveOptions](Word.WdSaveOptions.md)** constants: **wdDoNotSaveChanges**, **wdPromptToSaveChanges**, or **wdSaveChanges**.|
| _OriginalFormat_|Optional| **Variant**|Specifies the save format for the document. Can be one of the following **[WdOriginalFormat](Word.WdOriginalFormat.md)** constants: **wdOriginalDocumentFormat**, **wdPromptUser**, or **wdWordDocument**.|
| _RouteDocument_|Optional| **Variant**| **True** to route the document to the next recipient. If the document doesn't have a routing slip attached, this argument is ignored.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]