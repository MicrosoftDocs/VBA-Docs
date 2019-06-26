---
title: Document.XMLBeforeDelete event (Word)
keywords: vbawd10.chm4001009
f1_keywords:
- vbawd10.chm4001009
ms.prod: word
api_name:
- Word.Document.XMLBeforeDelete
ms.assetid: 1cef9cdb-a80a-8d38-9646-e3353f6c6923
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.XMLBeforeDelete event (Word)

Occurs when a user deletes an XML element from a document. If more than one element is deleted from the document at the same time (for example, when cutting and pasting XML), the event fires for each element that is deleted.


## Syntax

_expression_.**XMLBeforeDelete** (_DeletedRange_, _OldXMLNode_, _InUndoRedo_)

_expression_ A variable that represents a **[Document](Word.Document.md)** object.




## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeletedRange_|Required| **[Range](Word.Range.md)**|The contents of the XML element being deleted. If only an element is deleted and not associated text, the DeletedRange parameter will not exist and will, therefore, be set to  **Nothing**.|
| _OldXMLNode_|Required| **[XMLNode](Word.XMLNode.md)**|The node that is being deleted.|
| _InUndoRedo_|Required| **Boolean**| **True** indicates the action was performed using the **Undo** or **Redo** feature in Microsoft Word.|

## Remarks

Represents a **Document** object that has been declared by using the **WithEvents** keyword in a class module. For information about using events with a **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).

If the _InUndoRedo_ parameter is **True**, never change the XML in a document while the **XMLAfterInsert** and **XMLBeforeDelete** events are running.

If the _InUndoRedo_ parameter is  **False**, you can insert and delete the XML in the document&mdash;but be careful that the **XMLAfterInsert** and **XMLBeforeDelete** events will not try to cancel each other out, causing an infinite loop. You can prevent infinite loops by using a global **Boolean** variable and check for that at the beginning of the error handler, as shown in the following example.

```vb
Dim blnIsXMLDeleteRunning As Boolean 
 
Private Sub Document_XMLBeforeDelete(ByVal DeletedRange As Range, _ 
 ByVal OldXMLNode As XMLNode, ByVal InUndoRedo As Boolean) 
 
 If blnIsXMLDeleteRunning = False Then 
 blnIsXMLDeleteRunning = True 
 'Insert your event code here. 
 Else 
 Exit Sub 
 End If 
End Sub
```


## Example

The following example runs when an XML element is deleted. If the element contains text, a message is displayed asking whether the user wants to delete the text the element contains. If the user responds by clicking No, the contents of the element are copied to the Clipboard.


```vb
Private Sub Document_XMLBeforeDelete(ByVal DeletedRange As Range, _ 
 ByVal OldXMLNode As XMLNode, ByVal InUndoRedo As Boolean) 
 
 Dim intResponse As Integer 
 
 If InUndoRedo = False Then 
 If Not DeletedRange Is Nothing Then 
 intResponse = MsgBox("Are you sure you want to delete the text " _ 
 & vbCrLf & DeletedRange.Text, vbYesNo) 
 
 If intResponse = vbNo Then 
 
 DeletedRange.Copy 
 
 MsgBox "The text has been copied to the Clipboard." & vbCrLf & _ 
 "Position your cursor where you want to insert it, " & _ 
 vbCrLf & " and click Paste on the Edit menu." 
 
 End If 
 End If 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]