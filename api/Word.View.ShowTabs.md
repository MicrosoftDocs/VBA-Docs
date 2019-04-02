---
title: View.ShowTabs property (Word)
keywords: vbawd10.chm161808399
f1_keywords:
- vbawd10.chm161808399
ms.prod: word
api_name:
- Word.View.ShowTabs
ms.assetid: eca4147b-323f-10f3-e604-b3d9394bbbef
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowTabs property (Word)

 **True** if tab characters are displayed. Read/write **Boolean**.


## Syntax

_expression_. `ShowTabs`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example inserts a tab before the selection and displays tab characters in the window for Document2.


```vb
With Windows("Document2") 
 .Activate 
 .View.ShowTabs = True 
End With 
Selection.InsertBefore vbTab 
Selection.Collapse Direction:=wdCollapseEnd
```

This example splits the active window, shows tab characters in the first pane, and hides tab characters in the second pane.




```vb
With ActiveDocument.ActiveWindow 
 .Split = True 
 .Panes(1).View.ShowTabs = True 
 .Panes(2).View.ShowTabs = False 
End With
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]