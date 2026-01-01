---
title: Documents.AfterDocumentMerge event (Visio)
api_name:
- Visio.documents.afterdocumentmerge
ms.assetid: cac0544d-77b9-b722-cfdb-e42475ce2558
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Documents.AfterDocumentMerge event (Visio)

Occurs when Visio incorporates changes from other users' versions of a document into a merged, co authored document.


## Syntax

_expression_.**AfterDocumentMerge** (_coauthMergeObjects_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _coauthMergeObjects_|Required|**[IVCOAUTHMERGEEVENT]**|An object that represents different versions of the merged, co authored document.|

## See also


[Documents Object](Visio.Documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]