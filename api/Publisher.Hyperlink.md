---
title: Hyperlink Object (Publisher)
keywords: vbapb10.chm4653055
f1_keywords:
- vbapb10.chm4653055
ms.prod: publisher
api_name:
- Publisher.Hyperlink
ms.assetid: 1cc6d95b-357a-c169-a5d2-6850a1a3bbd6
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink Object (Publisher)

Represents a hyperlink. The  **Hyperlink** object is a member of the **[Hyperlinks](Publisher.Hyperlinks.md)** collection and the **[Shape](./Publisher.Shape.md)** and **[ShapeRange](Publisher.ShapeRange.md)** objects.


## Example

Use the  **[Hyperlink](./Publisher.Shape.Hyperlink.md)** property to return a **Hyperlink** object associated with a shape (a shape can have only one hyperlink). The following example deletes the hyperlink associated with the first shape in the active document.


```vb
Sub DeleteHyperlink() 
 ActiveDocument.Pages(1).Shapes(1).Hyperlink.Delete 
End Sub
```

Use  **Hyperlinks** (index), where index is the index number, to return a single **Hyperlink** object from a document, range, or selection. The following example deletes the first hyperlink in the selection.




```vb
Sub DeleteSelectedHyperlink() 
 If Selection.TextRange.Hyperlinks.Count >= 1 Then 
 Selection.TextRange.Hyperlinks(1).Delete 
 End If 
End Sub
```

Use the  **[Add](./Publisher.Hyperlinks.Add.md)** method to add a hyperlink. The following example adds a hyperlink to the selected text.




```vb
Sub AddHyperlinkToSelectedText() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="https://www.tailspintoys.com/" 
End Sub
```

Use the  **[Address](./Publisher.Hyperlink.Address.md)** property to add or change the address to a hyperlink. The following example adds a shape to the active publication and then adds a hyperlink to the shape.




```vb
Sub AddHyperlinkToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=200, Width:=300, Height:=300) 
 .Hyperlink.Address = "https://www.tailspintoys.com/" 
 End With 
End Sub
```


## Methods



|Name|
|:-----|
|[Delete](./Publisher.Hyperlink.Delete.md)|
|[SetPageRelative](./Publisher.Hyperlink.SetPageRelative.md)|

## Properties



|Name|
|:-----|
|[Address](./Publisher.Hyperlink.Address.md)|
|[Application](./Publisher.Hyperlink.Application.md)|
|[EmailSubject](./Publisher.Hyperlink.EmailSubject.md)|
|[PageID](./Publisher.Hyperlink.PageID.md)|
|[Parent](./Publisher.Hyperlink.Parent.md)|
|[Range](./Publisher.Hyperlink.Range.md)|
|[Shape](./Publisher.Hyperlink.Shape.md)|
|[TargetType](./Publisher.Hyperlink.TargetType.md)|
|[TextToDisplay](./Publisher.Hyperlink.TextToDisplay.md)|
|[Type](./Publisher.Hyperlink.Type.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]