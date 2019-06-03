---
title: Application.Documents property (Publisher)
keywords: vbapb10.chm131174
f1_keywords:
- vbapb10.chm131174
ms.prod: publisher
api_name:
- Publisher.Application.Documents
ms.assetid: dd48d68f-a6ae-b5c0-2a85-90abff1e6c5a
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.Documents property (Publisher)

Returns a **[Documents](Publisher.Documents.md)** collection that represents all open publications. Read-only.


## Syntax

_expression_.**Documents**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

Documents


## Example

The following example lists all the open publications.

```vb
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg & objDocument.Name & vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]