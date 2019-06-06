---
title: Document.BeginCustomUndoAction method (Publisher)
keywords: vbapb10.chm196709
f1_keywords:
- vbapb10.chm196709
ms.prod: publisher
api_name:
- Publisher.Document.BeginCustomUndoAction
ms.assetid: 316f443e-6782-594b-b955-f5ab60140f6a
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.BeginCustomUndoAction method (Publisher)

Specifies the starting point and label (textual description) of a group of actions that are wrapped to create a single undo action. 

The **[EndCustomUndoAction](Publisher.Document.EndCustomUndoAction.md)** method is used to specify the endpoint of the actions used to create the single undo action. The wrapped group of actions can be undone with a single undo.


## Syntax

_expression_.**BeginCustomUndoAction** (_ActionName_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ActionName_|Required| **String**|The label that corresponds to the single undo action. This label appears when you choose the arrow beside the **Undo** button on the **Standard** toolbar.|

## Remarks

The following methods of the **Document** object are unavailable within a custom undo action. A run-time error is returned if any of these methods are called within a custom undo action:

- **Close**    
- **MailMerge.DataSource.Close**    
- **PrintOut**    
- **Redo**    
- **Save**    
- **SaveAs**    
- **Undo**   
- **UndoClear**   
- **UpdateOLEObjects**
    
The **BeginCustomUndoAction** method must be called before the **EndCustomUndoAction** method is called. A run-time error is returned if **EndCustomUndoAction** is called before **BeginCustomUndoAction**.

Nesting a custom undo action within another custom undo action is allowed, but the nested custom undo action has no effect. Only the outermost custom undo action is active.


## Example

The following example contains two custom undo actions. The first one is created on the first page of the active publication. The **BeginCustomUndoAction** method is used to specify the point at which the custom undo action should begin. Six individual actions are performed, and then they are wrapped into one action with the call to **EndCustomUndoAction**. 

The text in the text frame that was created within the first custom undo action is then tested to determine whether the font is Verdana. If not, the **[Undo](Publisher.Document.Undo(method).md)** method is called with **[UndoActionsAvailable](Publisher.Document.UndoActionsAvailable.md)** passed as a parameter. In this case, there is only one undo action available. So, the call to the **Undo** method undoes only one action, but this one action has wrapped six actions into one.

A second undo action is then created, and it could also be undone later with a single undo operation.

```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(1) 
 
With theDoc 
 ' The following six actions are wrapped to create one 
 ' custom undo action named "Add Rectangle and Courier Text". 
 .BeginCustomUndoAction ("Add Rectangle and Courier Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 14 
 .Font.Bold = msoTrue 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 .EndCustomUndoAction 
 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 ' This call to Undo will undo all actions that are available. 
 ' In this case, there is only one action that can be undone. 
 .Undo (.UndoActionsAvailable) 
 ' A new custom undo action is created with a name of 
 ' "Add Balloon and Verdana Text". 
 .BeginCustomUndoAction ("Add Balloon and Verdana Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeBalloon, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 11 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End With 
 .EndCustomUndoAction 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]