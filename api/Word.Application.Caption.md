---
title: Application.Caption property (Word)
keywords: vbawd10.chm158335056
f1_keywords:
- vbawd10.chm158335056
ms.prod: word
api_name:
- Word.Application.Caption
ms.assetid: 5554fa04-0744-400d-fd8c-2fe36d4ad9a3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Caption property (Word)

Returns or sets the text displayed in the Title bar of the application window. Read/write  **String**.


## Syntax

_expression_.**Caption**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

To change the caption of the application window to the default text, set this property to an empty string ("").


## Example

This example resets the caption of the application window.


```vb
Application.Caption = ""
```

This example changes the caption of the Word application window to include the user name.




```vb
Application.Caption = UserName & "'s copy of Word"
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]