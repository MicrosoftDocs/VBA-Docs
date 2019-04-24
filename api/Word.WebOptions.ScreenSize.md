---
title: WebOptions.ScreenSize property (Word)
keywords: vbawd10.chm165937160
f1_keywords:
- vbawd10.chm165937160
ms.prod: word
api_name:
- Word.WebOptions.ScreenSize
ms.assetid: 4398a153-6932-17ef-b449-a532363fb428
ms.date: 06/08/2017
localization_priority: Normal
---


# WebOptions.ScreenSize property (Word)

Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a web browser. Read/write  **MsoScreenSize**.


## Syntax

_expression_.**ScreenSize**

_expression_ Required. A variable that represents a **[WebOptions](Word.WebOptions.md)** collection.


## Example

This example sets the target screen size for the active Web page at 800x600 pixels.


```vb
ActiveDocument.WebOptions.ScreenSize = _ 
 msoScreenSize800x600
```


## See also


[WebOptions Object](Word.WebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]