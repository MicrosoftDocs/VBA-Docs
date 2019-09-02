---
title: Documents.CheckOut method (Word)
keywords: vbawd10.chm158072848
f1_keywords:
- vbawd10.chm158072848
ms.prod: word
api_name:
- Word.Documents.CheckOut
ms.assetid: 70b89f66-7d02-ad40-d868-f6aa7b13ebdd
ms.date: 09/03/2019
localization_priority: Normal
---


# Documents.CheckOut method (Word)

Copies a specified document from a server to a local computer for editing.


## Syntax

_expression_.**CheckOut** (_FileName_)

_expression_ Required. A variable that represents a **[Documents](Word.Documents.md)** object.



## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to check out.|

## Remarks

To take advantage of the collaboration features built into Word, documents must be stored on a Microsoft SharePoint Portal Server.


## Example

This example verifies that a document is not checked out by another user and that it can be checked out. If the document can be checked out, it copies the document to the local computer for editing.

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

