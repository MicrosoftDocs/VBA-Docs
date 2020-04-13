---
title: Application.Quit event (Word)
keywords: vbawd10.chm400002
f1_keywords:
- vbawd10.chm400002
ms.prod: word
api_name:
- Word.Application.Quit
ms.assetid: 3e05cf42-47c9-6a1b-b7da-09abe9746fd5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Quit event (Word)

Occurs when the user exits Microsoft Word.


## Syntax

Private Sub Application_**Quit**()

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module.


## Remarks

For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Example

This example ensures that the **Standard** and **Formatting** toolbars are visible before the user exits Word. As a result, when Word is started again, the **Standard** and **Formatting** toolbars are visible.

This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md)for directions on how to accomplish this.




```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_Quit() 
 CommandBars("Standard").Visible = True 
 CommandBars("Formatting").Visible = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]