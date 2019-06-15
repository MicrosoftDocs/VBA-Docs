---
title: TabStop.Application property (Publisher)
keywords: vbapb10.chm5636097
f1_keywords:
- vbapb10.chm5636097
ms.prod: publisher
api_name:
- Publisher.TabStop.Application
ms.assetid: 189ea579-7859-1db7-c4c1-9616660d49d9
ms.date: 06/15/2019
localization_priority: Normal
---


# TabStop.Application property (Publisher)

When used without an object qualifier, this property returns an **[Application](Publisher.Application.md)** object that represents the current instance of Publisher. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. 

When used with an OLE Automation object, it returns the object's application.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[TabStop](Publisher.TabStop.md)** object.


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