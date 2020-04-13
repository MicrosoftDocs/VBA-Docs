---
title: HeaderFooter.Shapes property (Word)
keywords: vbawd10.chm159711239
f1_keywords:
- vbawd10.chm159711239
ms.prod: word
api_name:
- Word.HeaderFooter.Shapes
ms.assetid: dc38943b-b4fa-51c5-ff3d-8180ff51c279
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter.Shapes property (Word)

Returns a  **Shapes** collection that represents all the **Shape** objects in a header or footer. Read-only.


## Syntax

_expression_.**Shapes**

_expression_ A variable that represents a '[HeaderFooter](Word.HeaderFooter.md)' object.


## Remarks

This collection can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).

The **Shapes** property, when applied to a document, returns all the **Shape** objects in the main story of the document, excluding the headers and footers. When applied to a **HeaderFooter** object, the **Shapes** property returns all the **Shape** objects found in all the headers and footers in the document.


## Example

This example displays a count of all the shapes in the primary header and footer of the first section of the active document.


```vb
MsgBox ActiveDocument.Sections(1). _ 
 Headers(wdHeaderFooterPrimary).Shapes.Count
```


## See also


[HeaderFooter Object](Word.HeaderFooter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]