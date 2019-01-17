---
title: Application.Documents Property (Publisher)
keywords: vbapb10.chm131174
f1_keywords:
- vbapb10.chm131174
ms.prod: publisher
api_name:
- Publisher.Application.Documents
ms.assetid: dd48d68f-a6ae-b5c0-2a85-90abff1e6c5a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Documents Property (Publisher)

Returns a  **[Documents](Publisher.Documents.md)** collection that represents all open publications. Read-only.


## Syntax

 _expression_. **Documents**

 _expression_ A variable that represents a  **Application** object.


## Return value

Documents


## Example

The following example lists all of the open publications.


```vb
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg & objDocument.Name & vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]