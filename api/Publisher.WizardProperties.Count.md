---
title: WizardProperties.Count Property (Publisher)
keywords: vbapb10.chm1507331
f1_keywords:
- vbapb10.chm1507331
ms.prod: publisher
api_name:
- Publisher.WizardProperties.Count
ms.assetid: 835f3467-ec89-54d2-c685-3021e6267121
ms.date: 06/08/2017
---


# WizardProperties.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a  **WizardProperties** object.


## Example

This example displays the number of pages in the active document.


```vb
Sub CountNumberOfPages() 
<<<<<<< HEAD
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
=======
 MsgBox "Your publication contains " & _ 
 ActiveDocument.Pages.Count & " page(s)." 
>>>>>>> master
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
 
<<<<<<< HEAD
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
=======
 MsgBox "Your publication contains " & intShapes & " shape(s)." 
>>>>>>> master
End Sub
```


