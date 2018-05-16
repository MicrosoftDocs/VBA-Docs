---
title: OLEFormat Object (Publisher)
keywords: vbapb10.chm4521983
f1_keywords:
- vbapb10.chm4521983
ms.prod: publisher
api_name:
- Publisher.OLEFormat
ms.assetid: e5b72d6b-dff8-3882-549f-e376c1e4d372
ms.date: 06/08/2017
---


# OLEFormat Object (Publisher)

Represents the OLE characteristics, other than linking (see the  **[LinkFormat](Publisher.LinkFormat.md)** object), for an OLE object, ActiveX control, or field.
 


## Remarks

Not all types of shapes and fields have OLE capabilities. Use the  **[Type](Publisher.Shape.Type.md)** property for the **[Shape](Publisher.Shape.md)** object to determine into which category the specified shape falls.
 

 
Use the  **[Activate](Publisher.OLEFormat.Activate.md)** and **[DoVerb](Publisher.OLEFormat.DoVerb.md)** methods to automate an OLE object.
 

 

## Example

Use the  **[OLEFormat](Publisher.Shape.OLEFormat.md)** property for a shape or field to return an **OLEFormat** object. The following example activates all OLE objects in the active publication.
 

 

```
Sub ActivateOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.OLEFormat.Activate 
 End If 
 Next 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Activate](Publisher.OLEFormat.Activate.md)|
|[DoVerb](Publisher.OLEFormat.DoVerb.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.OLEFormat.Application.md)|
|[Object](Publisher.OLEFormat.Object.md)|
|[ObjectVerbs](Publisher.OLEFormat.ObjectVerbs.md)|
|[Parent](oleformat-parent-property-publisher.md)|
|[ProgId](oleformat-progid-property-publisher.md)|

