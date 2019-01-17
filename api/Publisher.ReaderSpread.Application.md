---
title: ReaderSpread.Application Property (Publisher)
keywords: vbapb10.chm524289
f1_keywords:
- vbapb10.chm524289
ms.prod: publisher
api_name:
- Publisher.ReaderSpread.Application
ms.assetid: 84bff0f2-db25-debd-25fb-bfb115600571
ms.date: 06/08/2017
localization_priority: Normal
---


# ReaderSpread.Application Property (Publisher)

Used without an object qualifier, this property returns an  **[Application](Publisher.Application.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a  **ReaderSpread** object.


## Example

This example displays the version and build information for Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 & .Version & " build " & .Build 
End With
```

This example displays the name of the application that created each linked OLE object on page one of the active publication.




```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


