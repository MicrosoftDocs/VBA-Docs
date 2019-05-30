---
title: BorderArt object (Publisher)
keywords: vbapb10.chm7733247
f1_keywords:
- vbapb10.chm7733247
ms.prod: publisher
api_name:
- Publisher.BorderArt
ms.assetid: 464bec0f-7912-ab27-9593-7f1cb53da342
ms.date: 05/31/2019
localization_priority: Normal
---


# BorderArt object (Publisher)

Represents an available type of BorderArt. BorderArt is picture borders that can be applied to text boxes, picture frames, or rectangles. The **BorderArt** object is a member of the **[BorderArts](Publisher.BorderArts.md)** collection. The **BorderArts** collection contains all BorderArt available for use in the specified publication.
 

## Remarks

The **BorderArts** collection includes any custom BorderArt types created by the user for the specified publication.

Use the **[Item](Publisher.BorderArts.Item.md)** property of the **BorderArts** collection to return a specific BorderArt object. The _Index_ argument of the **Item** property can be the number or name of the BorderArt object.

Use the **[Name](Publisher.BorderArt.Name.md)** property to specify which type of BorderArt you want applied to a picture.

> [!NOTE] 
> Because **Name** is the default property of both the **BorderArt** object and the **BorderArtFormat** object, you do not need to state it explicitly when setting the BorderArt type. The statement `Shape.BorderArtFormat = Document.BorderArts(1)` is equivalent to `Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`.

## Example

This example returns the BorderArt named Apples from the active publication. 

```vb
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
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


## Properties

- [Application](Publisher.BorderArt.Application.md)
- [Name](Publisher.BorderArt.Name.md)
- [Parent](Publisher.BorderArt.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]