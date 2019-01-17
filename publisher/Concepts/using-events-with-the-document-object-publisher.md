---
title: Using Events with the Document Object (Publisher)
ms.prod: publisher
ms.assetid: 0f5cfe67-bfa1-0ec7-11c9-c4c1337ebe50
ms.date: 06/08/2017
localization_priority: Normal
---


# Using Events with the Document Object (Publisher)

The  **Document** object supports seven events: **[BeforeClose](../../api/Publisher.Document.BeforeClose.md)**,  **[Open](../../api/Publisher.Document.Open.md)**,  **[Redo](../../api/Publisher.Document.Redo(even).md)**,  **[ShapesAdded](../../api/Publisher.Document.ShapesAdded.md)**,  **[ShapesRemoved](../../api/Publisher.Document.ShapesRemoved.md)**,  **[Undo](../../api/Publisher.Document.Undo(even).md)**, and  **[WizardAfterChange](../../api/Publisher.Document.WizardAfterChange.md)**. You write procedures to respond to these events in the class module named "ThisDocument". Use the following steps to create an event procedure.


1. Under your publication project in the  **Project Explorer** window, double-click **ThisDocument**. (In  **Folder** view, **ThisDocument** is located in the **Microsoft Publisher Objects** folder.)
    
2. Select  **Document** from the **Object** drop-down list box.
    
3. Select an event from the  **Procedure** drop-down list box. An empty subroutine is added to the class module.
    
4. Add the Visual Basic instructions you want to run when the event occurs.
    

## Example

This example shows an  **Open** event procedure that displays a message when a publication is opened.


```vb
Private Sub Document_Open() 
    MsgBox "This publication is copyrighted." 
End Sub
```

The following example shows a  **BeforeClose** event procedure that prompts the user for a yes or no response before closing a document.




```vb
Private Sub Document_BeforeClose(Cancel As Boolean) 
    Dim intResponse As Integer 
 
    intResponse = MsgBox("Do you really want to close " _ 
        & "the document?", vbYesNo) 
 
    If intResponse = vbNo Then Cancel = True 
End Sub
```


 **Note**  For information on creating event procedures for the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md) .

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]