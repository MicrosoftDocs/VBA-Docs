---
title: Row.Count property (Visio)
keywords: vis_sdr.chm15813330
f1_keywords:
- vis_sdr.chm15813330
ms.prod: visio
api_name:
- Visio.Row.Count
ms.assetid: 0fb1c315-3e65-3397-c047-2a59dfaa4e31
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Count property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[Row](Visio.Row.md)** object.


## Return value

Integer


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Count** property to iterate through a **Documents** collection. It displays the names of all the open Microsoft Visio documents in the Immediate window.


```vb
 
Public Sub Count_Example() 
 
 Dim intCounter As Integer 
 Dim vsoDocument As Visio.Document 
 
 For intCounter = 1 To Documents.Count 
 'Get the next open document. 
 Set vsoDocument = Documents.Item(intCounter) 
 
 'Print its name in the Immediate window. 
 Debug.Print vsoDocument.Name 
 Next intCounter 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]