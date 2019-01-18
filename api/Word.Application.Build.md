---
title: Application.Build property (Word)
keywords: vbawd10.chm158335023
f1_keywords:
- vbawd10.chm158335023
ms.prod: word
api_name:
- Word.Application.Build
ms.assetid: e22e7633-9327-eacc-3936-3d113381f675
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Build property (Word)

Returns the version and build number of the Word application. Read-only  **String**.


## Syntax

 _expression_. `Build`

 _expression_ A variable that represents an '[Application](Word.Application.md)' object.


## Example

This example displays the version and build number of Word.


```vb
MsgBox Prompt:=Application.Build, _ 
 Title:="Microsoft Word Version"
```


## See also


[Application Object](Word.Application.md)

