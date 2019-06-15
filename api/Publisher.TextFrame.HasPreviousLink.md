---
title: TextFrame.HasPreviousLink property (Publisher)
keywords: vbapb10.chm3866641
f1_keywords:
- vbapb10.chm3866641
ms.prod: publisher
api_name:
- Publisher.TextFrame.HasPreviousLink
ms.assetid: 85e0b497-55c9-d49f-2b65-e199361c121a
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.HasPreviousLink property (Publisher)

Returns **msoTrue** if the specified text frame has a valid link to a backward text box, and returns **msoFalse** if it does not. Read-only.


## Syntax

_expression_.**HasPreviousLink**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Return value

**[MsoTriState](office.msotristate.md)**


## Example

This example breaks all links in the document to the first specified text frame if links exist. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub AddPreviousNextLinkPages() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 If .HasNextLink Then .BreakForwardLink 
 If .HasPreviousLink Then .PreviousLinkedTextFrame _ 
 .BreakForwardLink 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]