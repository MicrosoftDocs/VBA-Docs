---
title: Tags object (Publisher)
keywords: vbapb10.chm4718591
f1_keywords:
- vbapb10.chm4718591
ms.prod: publisher
api_name:
- Publisher.Tags
ms.assetid: 76cccc1e-4623-af8b-f0f8-e6cc245b94fd
ms.date: 06/01/2019
localization_priority: Normal
---


# Tags object (Publisher)

A collection of **[Tag](Publisher.Tag.md)** objects that represent tags or custom properties applied to a shape, shape range, page, or publication.
 
## Remarks

Use the **[Shape.Tags](Publisher.Shape.Tags.md)** property to access the **Tags** collection. 

Use the **Add** method to add a **Tag** object to a shape, shape range, page, or publication. 

Use the **Count** property to determine if a shape, shape range, page, or publication contains one or more **Tag** objects. 


## Example

This example adds a tag to each oval shape on the first page of the active publication.

```vb
Sub AddNewTag() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If InStr(1, shp.Name, "Oval") > 0 Then 
 shp.Tags.Add Name:="Shape", Value:="Oval" 
 End If 
 Next shp 
 End With 
End Sub
```

<br/>

This example fills all shapes on the first page of the active publication if the shape's first tag has a value of Oval.

```vb
Sub FormatTaggedShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.Tags.Count > 0 Then 
 If shp.Tags(1).Value = "Oval" Then 
 shp.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End If 
 End If 
 Next shp 
 End With 
End Sub
```


## Methods

- [Add](Publisher.Tags.Add.md)
- [Item](Publisher.Tags.Item.md)

## Properties

- [Application](Publisher.Tags.Application.md)
- [Count](Publisher.Tags.Count.md)
- [Parent](Publisher.Tags.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]