---
title: BorderArtFormat object (Publisher)
keywords: vbapb10.chm7667711
f1_keywords:
- vbapb10.chm7667711
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat
ms.assetid: ba066b2e-fe40-aeef-9275-2cc2810f63ca
ms.date: 05/31/2019
localization_priority: Normal
---


# BorderArtFormat object (Publisher)

Represents the formatting of the BorderArt applied to the specified shape.
 
## Remarks

BorderArt are picture borders that can be applied to text boxes, picture frames, or rectangles.
 
Use the **[BorderArt](Publisher.Shape.BorderArt.md)** property of a shape to return a **BorderArtFormat** object.

Use the **Set** method to specify which type of BorderArt you want applied to a picture. 

You can also use the **Name** property to specify which type of BorderArt you want applied to a picture. 

> [!NOTE] 
> Because **Name** is the default property of both the **[BorderArt](Publisher.BorderArt.md)** and **BorderArtFormat** objects, you do not need to state it explicitly when setting the BorderArt type. The statement `Shape.BorderArtFormat = Document.BorderArts(1)` is equivalent to `Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`.
 
Use the **Delete** method to remove BorderArt from a picture. 
 
## Example

The following example returns the BorderArt of the first shape on the first page of the active publication, and displays the name of the BorderArt in a message box.

```vb
Dim bdaTemp As BorderArtFormat 
 
Set bdaTemp = ActiveDocument.Pages(1).Shapes(1).BorderArt 
MsgBox "BorderArt name is: " &bdaTemp.Name
```

<br/>

The following example tests for the existence of BorderArt on each shape for each page of the active document. Any BorderArt found is set to the same type.

```vb
Sub SetBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Set(strBorderArtName) 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```

<br/>

The following example sets all the BorderArt in a document to the same type by using the **Name** property.

```vb
Sub SetBorderArtByName() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Name = strBorderArtName 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```

<br/>

The following example tests for the existence of border art on each shape for each page of the active document. If border art exists, it is deleted.

```vb
Sub DeleteBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Delete 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```


## Methods

- [Delete](Publisher.BorderArtFormat.Delete.md)
- [RevertToDefaultWeight](Publisher.BorderArtFormat.RevertToDefaultWeight.md)
- [RevertToOriginalColor](Publisher.BorderArtFormat.RevertToOriginalColor.md)
- [Set](Publisher.BorderArtFormat.Set.md)

## Properties

- [Application](Publisher.BorderArtFormat.Application.md)
- [Color](Publisher.BorderArtFormat.Color.md)
- [Exists](Publisher.BorderArtFormat.Exists.md)
- [Name](Publisher.BorderArtFormat.Name.md)
- [Parent](Publisher.BorderArtFormat.Parent.md)
- [StretchPictures](Publisher.BorderArtFormat.StretchPictures.md)
- [Weight](Publisher.BorderArtFormat.Weight.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]