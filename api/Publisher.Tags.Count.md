---
title: Tags.Count Property (Publisher)
keywords: vbapb10.chm4653059
f1_keywords:
- vbapb10.chm4653059
ms.prod: publisher
api_name:
- Publisher.Tags.Count
ms.assetid: 46d443a3-643b-a43f-a77e-19a32af67217
ms.date: 06/08/2017
localization_priority: Normal
---


# Tags.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a  **Tags** object.


## Example

This example displays the number of pages in the active document.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " & _ 
 ActiveDocument.Pages.Count & " page(s)." 
End Sub
```

This example displays the number of shapes in the active document.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " & intShapes & " shape(s)." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]