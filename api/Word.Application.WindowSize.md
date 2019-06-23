---
title: Application.WindowSize event (Word)
keywords: vbawd10.chm4000024
f1_keywords:
- vbawd10.chm4000024
ms.prod: word
api_name:
- Word.Application.WindowSize
ms.assetid: 96d55786-52c8-68a9-b9e9-b29c320a435a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowSize event (Word)

Occurs when the application window is resized or moved.


## Syntax

_expression_.**WindowSize** (_Doc_, _Wn_)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module. For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document in the window being sized.|
| _Wn_|Required| **Window**|The window being sized.|

## Example

This example displays a message every time the Microsoft Word application window is moved or resized. This example assumes that you have declared an application variable called "WordApp" in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub WordApp_WindowSize(ByVal Doc As Document, _ 
 ByVal Wn As Window) 
 MsgBox "You have just resized or moved your window." 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]