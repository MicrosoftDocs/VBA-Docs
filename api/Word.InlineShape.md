---
title: InlineShape object (Word)
keywords: vbawd10.chm2472
f1_keywords:
- vbawd10.chm2472
ms.prod: word
api_name:
- Word.InlineShape
ms.assetid: a8fd110a-4aa7-c4b9-1559-32022787d955
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape object (Word)

Represents an object in the text layer of a document. An inline shape can only be a picture, an OLE object, or an ActiveX control. The **InlineShape** object is a member of the **[InlineShapes](Word.inlineshapes.md)** collection. The **InlineShapes** collection contains all the shapes that appear inline in a document, range, or selection.


## Remarks

 **InlineShape** objects are treated like characters and are positioned as characters within a line of text.

Use  **InlineShapes** (Index), where Index is the index number, to return a single **InlineShape** object. Inline shapes don't have names. The following example activates the first inline shape in the active document.




```vb
ActiveDocument.InlineShapes(1).Activate
```

 **Shape** objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. You can use the **ConvertToInlineShape** method and the **ConvertToShape** method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes. Use the **Type** property to return the type of inline shape: picture, linked picture, embedded OLE object, linked OLE object, or ActiveX control.


> [!NOTE] 
> When you open a document created in an earlier version of Word, pictures are converted to inline shapes.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
