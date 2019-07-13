---
title: Document.SolutionXMLElementExists property (Visio)
keywords: vis_sdr.chm10550865
f1_keywords:
- vis_sdr.chm10550865
ms.prod: visio
api_name:
- Visio.Document.SolutionXMLElementExists
ms.assetid: d4a0bd9b-a3ea-de0a-5c33-ccad4d4398eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SolutionXMLElementExists property (Visio)

Indicates whether a named SolutionXML element exists in the document. Read-only.


## Syntax

_expression_.**SolutionXMLElementExists** (_ElementName_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ElementName_|Required| **String**|The case-sensitive name of the SolutionXML element.|

## Return value

Boolean


## Remarks

Because the **SolutionXMLElement** property can overwrite existing XML data, always use the **SolutionXMLElementExists** property to verify whether _ElementName_ already exists in the document.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]