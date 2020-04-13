---
title: Application.FileValidation property (Word)
keywords: vbawd10.chm158335469
f1_keywords:
- vbawd10.chm158335469
ms.prod: word
api_name:
- Word.Application.FileValidation
ms.assetid: 2f88d1a7-98a7-9ec6-09b3-a09c1a934e01
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FileValidation property (Word)

Returns or sets how Word will validate files before opening them. Read/write [MsoFileValidationMode](Office.MsoFileValidationMode.md).


## Syntax

_expression_. `FileValidation`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Return value

[MsoFileValidationMode](Office.MsoFileValidationMode.md)


## Remarks

Files that do not pass validation will be opened in a [Protected View window](Word.ProtectedViewWindow.md). The **FileValidation** property is per session only. If you set the **FileValidation** property, that setting will remain in effect for the entire session the application is open.


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]