---
title: LinkFormat object (Word)
keywords: vbawd10.chm2353
f1_keywords:
- vbawd10.chm2353
ms.prod: word
api_name:
- Word.LinkFormat
ms.assetid: ca37d4e2-e978-8e6a-1e7a-7e43cf41e6c2
ms.date: 06/08/2017
localization_priority: Normal
---


# LinkFormat object (Word)

Represents the linking characteristics for an OLE object or picture.


## Remarks

Use the **LinkFormat** property for a shape, inline shape, or field to return the **LinkFormat** object. The following example breaks the link for the first shape on the active document.


```vb
ActiveDocument.Shapes(1).LinkFormat.BreakLink
```

Not all types of shapes, inline shapes, and fields can be linked to a source. Use the **Type** property for the **Shape** and **InlineShape** objects to determine whether a particular shape can be linked. The **Type** property for a **Field** object returns the type of field.

You can use both the **Update** method and the **AutoUpdate** property to update links. To return or set the full path for a particular link's source file, use the **SourceFullName** property.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]