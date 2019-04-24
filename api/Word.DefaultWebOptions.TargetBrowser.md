---
title: DefaultWebOptions.TargetBrowser property (Word)
keywords: vbawd10.chm165871633
f1_keywords:
- vbawd10.chm165871633
ms.prod: word
api_name:
- Word.DefaultWebOptions.TargetBrowser
ms.assetid: e5d31e0c-d669-4b16-bf8d-0c5353732b17
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.TargetBrowser property (Word)

Sets or returns an  **MsoTargetBrowser** constant representing the target browser for documents viewed in a web browser. Read/write.


## Syntax

_expression_.**TargetBrowser**

_expression_ Required. A variable that represents a **[DefaultWebOptions](Word.DefaultWebOptions.md)** collection.


## Example

This example sets the target browser for all documents to Internet Explorer 6.


```vb
Sub GlobalTargetBrowser() 
 Application.DefaultWebOptions _ 
 .TargetBrowser = msoTargetBrowserIE6 
End Sub
```


## See also


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]