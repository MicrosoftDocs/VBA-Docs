---
title: Field.Application property (Publisher)
keywords: vbapb10.chm6094849
f1_keywords:
- vbapb10.chm6094849
ms.prod: publisher
api_name:
- Publisher.Field.Application
ms.assetid: 3250d161-b88c-8e2e-34e1-65ca42368bd6
ms.date: 06/07/2019
localization_priority: Normal
---


# Field.Application property (Publisher)

When used without an object qualifier, this property returns an **[Application](Publisher.Application.md)** object that represents the current instance of Publisher. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. 

When used with an OLE Automation object, it returns the object's application.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Field](Publisher.Field.md)** object.


## Example

This example displays the version and build information for Publisher.

```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 & .Version & " build " & .Build 
End With
```

<br/>

This example displays the name of the application that created each linked OLE object on page one of the active publication.

```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]