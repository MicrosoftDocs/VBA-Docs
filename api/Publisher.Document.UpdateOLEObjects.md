---
title: Document.UpdateOLEObjects Method (Publisher)
keywords: vbapb10.chm196706
f1_keywords:
- vbapb10.chm196706
ms.prod: publisher
api_name:
- Publisher.Document.UpdateOLEObjects
ms.assetid: 2c07e755-6f5c-5fd8-091c-fbe3bfae6692
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.UpdateOLEObjects Method (Publisher)

Updates linked and embedded OLE objects.


## Syntax

 _expression_. **UpdateOLEObjects**

 _expression_ A variable that represents a  **Document** object.


## Example

This example updates all OLE objects in the active publication.


```vb
Sub SearchAndUpdateOLEObjects() 
 ActiveDocument.UpdateOLEObjects 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]