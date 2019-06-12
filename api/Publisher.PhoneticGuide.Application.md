---
title: PhoneticGuide.Application property (Publisher)
keywords: vbapb10.chm6160385
f1_keywords:
- vbapb10.chm6160385
ms.prod: publisher
api_name:
- Publisher.PhoneticGuide.Application
ms.assetid: 27a0aca5-4033-bc46-4831-10e5e6061e4b
ms.date: 06/12/2019
localization_priority: Normal
---


# PhoneticGuide.Application property (Publisher)

When used without an object qualifier, this property returns an **[Application](Publisher.Application.md)** object that represents the current instance of Publisher. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. 

When used with an OLE Automation object, it returns the object's application.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[PhoneticGuide](Publisher.PhoneticGuide.md)** object.


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