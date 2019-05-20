---
title: Source.XML property (Word)
keywords: vbawd10.chm140836969
f1_keywords:
- vbawd10.chm140836969
ms.prod: word
api_name:
- Word.Source.XML
ms.assetid: 811cd77f-558d-a884-3ef3-911c79410b2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Source.XML property (Word)

Returns a  **String** that represents the XML markup for a **Source** object. Read-only.


## Syntax

_expression_.**XML**

 _expression_ An expression that returns a [Source](./Word.Source.md) object.


## Remarks

The Data parameter of the  **[Add](Word.Sources.Add.md)** method for the **[Sources](Word.Sources.md)** object requires a valid XML string that contains data for the necessary elements. Depending on the type of source that you add, the XML may change. To determine what the correct XML needs to be for a specific type of source, create a source through the **Create Source** dialog box, and then use the **XML** property to return the correct XML syntax.


## See also


[Source Object](Word.Source.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]