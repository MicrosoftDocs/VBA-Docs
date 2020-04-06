---
title: Application.WindowBeforeRightClick event (Word)
keywords: vbawd10.chm4000012
f1_keywords:
- vbawd10.chm4000012
ms.prod: word
api_name:
- Word.Application.WindowBeforeRightClick
ms.assetid: c2d550e5-6781-a05f-41f6-eb9839aef208
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowBeforeRightClick event (Word)

Occurs when the editing area of a document window is right-clicked, before the default right-click action.


## Syntax

Private Sub  _expression_._'WindowBeforeRightClick** (_Sel As Selection_**, **_Cancel As Boolean_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sel_|Required| **Selection**|The current selection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the default right-click action does not occur when the procedure is finished.|

## Remarks

For more information about using events with the  **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).

## Example

This example prompts the user for a yes or no response before executing the default right-click action. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowBeforeRightClick _ 
 (ByVal Sel As Selection, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Selection = " & Sel & vbLf & vbLf _ 
 & "Continue with operation on this selection?", _ 
 vbYesNo) 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]