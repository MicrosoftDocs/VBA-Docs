---
title: LinkFormat object (Publisher)
keywords: vbapb10.chm4456447
f1_keywords:
- vbapb10.chm4456447
ms.prod: publisher
api_name:
- Publisher.LinkFormat
ms.assetid: 5b588edd-b026-cfc7-4acb-77290ae4d297
ms.date: 05/31/2019
localization_priority: Normal
---


# LinkFormat object (Publisher)

Represents the linking characteristics for an OLE object or picture.
 

## Remarks

Not all types of shapes and fields can be linked to a source. Use the **[Type](Publisher.Shape.Type.md)** property of the **Shape** object to determine whether a particular shape can be linked.
 
Use the **Update** method to update links. To return or set the full path for a particular link's source file, use the **SourceFullName** property.
 
Use the **[Shape.LinkFormat](Publisher.Shape.LinkFormat.md)** property for a shape or field to return a **LinkFormat** object. 
 

## Example

The following example updates the links to all linked OLE objects on the first page of the active publication.

```vb
Sub FindOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.LinkFormat.Update 
 End If 
 Next shpShape 
End Sub
```


## Methods

- [Update](Publisher.LinkFormat.Update.md)

## Properties

- [Application](Publisher.LinkFormat.Application.md)
- [Parent](Publisher.LinkFormat.Parent.md)
- [SourceFullName](Publisher.LinkFormat.SourceFullName.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]