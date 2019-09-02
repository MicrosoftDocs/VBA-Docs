---
title: Documents.CanCheckOut method (Word)
keywords: vbawd10.chm158072849
f1_keywords:
- vbawd10.chm158072849
ms.prod: word
api_name:
- Word.Documents.CanCheckOut
ms.assetid: eaa052ff-0194-4c3f-a8e3-5a18ae77038e
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.CanCheckOut method (Word)

**True** if Microsoft Word can check out a specified document from a server. Read/write **Boolean**.


## Syntax

_expression_.**CanCheckOut** (_FileName_)

_expression_ Required. A variable that represents a **[Documents](Word.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The server path and name of the document.|

## Return value

Boolean


## Remarks

To take advantage of the collaboration features built into Word, documents must be stored on a Microsoft SharePoint Portal Server.


## Example

This example verifies that a document is not being edited by another user and that it can be checked out. If the document can be checked out, it copies the document to the local computer for editing.


```vb
Sub CheckInOut(docCheckOut As String) 
 If Documents.CanCheckOut(docCheckOut) = True Then 
 Documents.CheckOut docCheckOut 
 Else 
 MsgBox "You are unable to check out this document at this time." 
 End If 
End Sub
```

<br/>

To call the CheckInOut subroutine, use the following subroutine and replace the "https://servername/workspace/report.doc" file name with an actual file located on a server mentioned in the Remarks section.


```vb
Sub CheckDocInOut() 
 Call CheckInOut (docCheckIn:="https://servername/workspace/report.doc") 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]