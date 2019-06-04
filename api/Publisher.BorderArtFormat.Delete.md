---
title: BorderArtFormat.Delete method (Publisher)
keywords: vbapb10.chm7602184
f1_keywords:
- vbapb10.chm7602184
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Delete
ms.assetid: 3ec0576f-8304-2647-7309-b014b586c1b6
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.Delete method (Publisher)

Deletes the specified object.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Remarks

A run-time error occurs if the specified object does not exist.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active publication. If BorderArt exists, it is deleted.

```vb
Sub DeleteBorderArt() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Delete 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]