---
title: Application.Quit method (Word)
keywords: vbawd10.chm158336081
f1_keywords:
- vbawd10.chm158336081
ms.prod: word
api_name:
- Word.Application.Quit
ms.assetid: 0279d848-a8b7-dac7-1e84-a55c72789e3b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Quit method (Word)

Quits Microsoft Word and optionally saves or routes the open documents.


## Syntax

_expression_.**Quit** (_SaveChanges_, _OriginalFormat_, _RouteDocument_)

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|Specifies whether Word saves changed documents before closing. Can be one of the **WdSaveOptions** constants.|
| _OriginalFormat_|Optional| **Variant**|Specifies the way Word saves documents whose original format was not Word Document format. Can be one of the **WdOriginalFormat** constants.|
| _RouteDocument_|Optional| **Variant**| **True** to route the document to the next recipient. If the document does not have a routing slip attached, this argument is ignored.|

## Example

This example closes Word and prompts the user to save each document that has changed since it was last saved.


```vb
Application.Quit SaveChanges:=wdPromptToSaveChanges
```

This example prompts the user to save all documents. If the user clicks Yes, all documents are saved in the Word format before Word closes.




```vb
Dim intResponse As Integer 
 
intResponse = _ 
 MsgBox("Do you want to save all documents?", vbYesNo) 
If intResponse = vbYes Then Application.Quit _ 
 SaveChanges:=wdSaveChanges, OriginalFormat:=wdWordDocument
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
