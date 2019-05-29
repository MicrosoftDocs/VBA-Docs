---
title: Template.CustomDocumentProperties property (Word)
keywords: vbawd10.chm157941769
f1_keywords:
- vbawd10.chm157941769
ms.prod: word
api_name:
- Word.Template.CustomDocumentProperties
ms.assetid: b11e705f-7042-014c-4efc-2d2fba135ab2
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.CustomDocumentProperties property (Word)

Returns a **[DocumentProperties](Office.DocumentProperties.md)** collection that represents all the custom document properties for the specified document.


## Syntax

_expression_.**CustomDocumentProperties**

_expression_ Required. A variable that represents a **[Template](Word.Template.md)** object.


## Remarks

Use the **[BuiltInDocumentProperties](word.document.builtindocumentproperties.md)** property to return the collection of built-in document properties. Properties of type **msoPropertyTypeString** (**[MsoDocProperties](office.msodocproperties.md)**) cannot exceed 255 characters in length.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]