---
title: Document.DocumentDirection Property (Publisher)
keywords: vbapb10.chm196648
f1_keywords:
- vbapb10.chm196648
ms.prod: publisher
api_name:
- Publisher.Document.DocumentDirection
ms.assetid: b28961ad-7adc-3920-0e67-88bb53310d9b
ms.date: 06/08/2017
---


# Document.DocumentDirection Property (Publisher)

Returns or sets a  **PbDirectionType** constant that indicates whether text in the document is read from left to right or from right to left. Read/write.


## Syntax

 _expression_. **DocumentDirection**

 _expression_ A variable that represents a  **Document** object.


### Return value

PbDirectionType


## Remarks

The  **DocumentDirection** property value can be one of the **[PbDirectionType](Publisher.PbDirectionType.md)** constants declared in the Microsoft Publisher type library.

The  **DocumentDirection** property affects the way the document is read but not the flow of text in the document. For example, if the document has a binding edge and is printed on both sides of the page, the binding edge for a left-to-right document would be different from the binding edge of a right-to-left document.

To format the direction of text flow, use the  **[DefaultTextFlowDirection](Publisher.Options.DefaultTextFlowDirection.md)** property to specify the default text flow for the entire document, or use the  **[Orientation](Publisher.TextFrame.Orientation.md)** property for an individual text frame to specify a text flow direction other than the default for the specified text frame only.


## Example

This example sets the active publication to read from left to right.


```vb
Sub SetBiDiText() 
 ActiveDocument.DocumentDirection = pbDirectionRightToLeft 
End Sub
```


