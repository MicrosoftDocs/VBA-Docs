---
title: Document.SolutionXMLElementName property (Visio)
keywords: vis_sdr.chm10550870
f1_keywords:
- vis_sdr.chm10550870
ms.prod: visio
api_name:
- Visio.Document.SolutionXMLElementName
ms.assetid: 460993bc-090c-00ad-805f-ae4af832ceba
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SolutionXMLElementName property (Visio)

Returns the name of the SolutionXML element. Read-only.


## Syntax

_expression_.**SolutionXMLElementName** (_Index_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index of the SolutionXML element in the document.|

## Return value

String


## Remarks

The only way to retrieve SolutionXML data is by name. You can use the **SolutionXMLElementName** property to get the element name to pass to the **SolutionXMLElement** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]