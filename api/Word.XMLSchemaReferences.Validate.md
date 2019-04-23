---
title: XMLSchemaReferences.Validate method (Word)
keywords: vbawd10.chm116129892
f1_keywords:
- vbawd10.chm116129892
ms.prod: word
api_name:
- Word.XMLSchemaReferences.Validate
ms.assetid: 66e4ea2d-e26c-be4c-fe1d-d240449f30f3
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLSchemaReferences.Validate method (Word)

Validates all the XML schemas that are attached to a document.


## Syntax

_expression_. `Validate`

 _expression_ An expression that returns an '[XMLSchemaReferences](Word.XMLSchemaReferences.md)' object.


## Return value

Nothing


## Remarks

When you run the  **Validate** method, Microsoft Word populates the **[XMLSchemaViolations](overview/Word.md)** property of the **[Document](Word.Document.md)** object with a collection of the XML nodes that have validation errors.


## See also


[XMLSchemaReferences Collection](Word.XMLSchemaReferences.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]