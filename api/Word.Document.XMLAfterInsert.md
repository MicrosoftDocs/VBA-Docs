---
title: Document.XMLAfterInsert event (Word)
keywords: vbawd10.chm4001008
f1_keywords:
- vbawd10.chm4001008
ms.prod: word
api_name:
- Word.Document.XMLAfterInsert
ms.assetid: 6858fd64-e96b-308e-53eb-d79595fabac7
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.XMLAfterInsert event (Word)

Occurs when a user adds a new XML element to a document. If more than one element is added to the document at the same time (for example, when cutting and pasting XML), the event fires for each element that is inserted.


## Syntax

_expression_.**XMLAfterInsert'(**_NewXMLNode_**, **_InUndoRedo_**)

_expression_ A variable that represents a '[Document](Word.Document.md)' object that has been declared by using the **WithEvents** keyword in a class module. For information about using events with a **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewXMLNode_|Required| **[XMLNode](Word.XMLNode.md)**|The newly added XML node.|
| _InUndoRedo_|Required| **Boolean**| **True** indicates the action was performed using the **Undo** or **Redo** feature in Microsoft Word.|

## Remarks

If the InUndoRedo parameter is  **True**, never change the XML in a document while the **XMLAfterInsert** and **XMLBeforeDelete** events are running.

If the InUndoRedo parameter is  **False**, you can insert and delete the XML in the document, but be careful that the **XMLAfterInsert** and **XMLBeforeDelete** events will not try to cancel each other out, causing an infinite loop. You can prevent infinite loops by using a global **Boolean** variable and check for that at the beginning of the error handler, as shown in the following example.




```vb
Dim blnIsXMLInsertRunning As Boolean 
 
Private Sub Document_XMLAfterInsert(ByVal DeletedRange As Range, _ 
 ByVal OldXMLNode As XMLNode, ByVal InUndoRedo As Boolean) 
 
 If blnIsXMLInsertRunning = False Then 
 blnIsXMLInsertRunning = True 
 'Insert your event code here. 
 Else 
 Exit Sub 
 End If 
End Sub
```


## Example

The following example validates a newly added node and if the node is not valid, displays a message describing the validation error.


```vb
Private Sub Document_XMLAfterInsert(ByVal NewXMLNode As XMLNode, _ 
 ByVal InUndoRedo As Boolean) 
 
 NewXMLNode.Validate 
 
 If NewXMLNode.ValidationStatus <> wdXMLValidationStatusOK Then 
 MsgBox NewXMLNode.ValidationErrorText 
 End If 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]