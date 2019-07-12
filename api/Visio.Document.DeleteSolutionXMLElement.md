---
title: Document.DeleteSolutionXMLElement method (Visio)
keywords: vis_sdr.chm10550555
f1_keywords:
- vis_sdr.chm10550555
ms.prod: visio
api_name:
- Visio.Document.DeleteSolutionXMLElement
ms.assetid: 2f00680e-56b1-c99b-2739-9d331965f802
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.DeleteSolutionXMLElement method (Visio)

Deletes the named SolutionXML element.


## Syntax

_expression_.**DeleteSolutionXMLElement** (_ElementName_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ElementName_|Required| **String**|The case-sensitive name of the SolutionXML data element.|

## Return value

Nothing


## Remarks

The  _ElementName_ parameter is case-sensitive and should match the name passed as a parameter to the **SolutionXMLElement** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]