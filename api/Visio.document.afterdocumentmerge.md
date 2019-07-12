---
title: Document.AfterDocumentMerge event (Visio)
ms.prod: visio
ms.assetid: 50658da5-592a-4d16-908f-c6abe3050f09
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.AfterDocumentMerge event (Visio)

Occurs when Visio incorporates changes from other users' versions of a document into a merged, co authored document.


## Syntax

_expression_.**AfterDocumentMerge** (_coauthMergeObjects_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _coauthMergeObjects_|Required|**[IVCOAUTHMERGEEVENT]**|An object that represents different versions of the merged, co authored document.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]