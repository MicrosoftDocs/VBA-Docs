---
title: Document.Sync event (Word)
keywords: vbawd10.chm4001007
f1_keywords:
- vbawd10.chm4001007
ms.prod: word
api_name:
- Word.Document.Sync
ms.assetid: cc46cfdf-ae26-9bba-7084-64349859d304
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Sync event (Word)

> [!NOTE] 
> This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## Syntax

_expression_.**Sync'(**_SyncEventType_**)

_expression_ A variable that represents a '[Document](Word.Document.md)' object that has been declared using the **WithEvents** keyword in a class module. For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SyncEventType_|Required| **MsoSyncEventType**|The status of the document synchronization.|

## Example

The following example displays a message if the synchronization of a document in a Document Workspace fails.


```vb
Private Sub Document_Sync(ByVal SyncEventType As Office.MsoSyncEventType) 
 
 If SyncEventType = msoSyncEventDownloadFailed Or _ 
 SyncEventType = msoSyncEventUploadFailed Then 
 
 MsgBox "Document synchronization failed. " & _ 
 "Please contact your administrator " & vbCrLf & _ 
 "or try again later." 
 
 End If 
 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]