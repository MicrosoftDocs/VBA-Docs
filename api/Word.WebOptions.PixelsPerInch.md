---
title: WebOptions.PixelsPerInch property (Word)
keywords: vbawd10.chm165937161
f1_keywords:
- vbawd10.chm165937161
ms.prod: word
api_name:
- Word.WebOptions.PixelsPerInch
ms.assetid: b5f8db0d-b3f9-4834-8228-1b2ad1b8e180
ms.date: 06/08/2017
localization_priority: Normal
---


# WebOptions.PixelsPerInch property (Word)

Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. Read/write  **Long**.


## Syntax

_expression_.**PixelsPerInch**

_expression_ Required. A variable that represents a **[WebOptions](Word.WebOptions.md)** collection.


## Remarks

The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. 

This property determines the size of the images and cells on the specified Web page relative to the size of text whenever you view the saved document in a web browser. The physical dimensions of the resulting image or cell are the result of the original dimensions (in inches) multiplied by the number of pixels per inch.

Use the  **[ScreenSize](Word.WebOptions.ScreenSize.md)** property to set the optimum screen size for the targeted web browsers.


## Example

This example sets the pixel density depending on the target screen size of the web browser.


```vb
With ActiveDocument.WebOptions 
 Select Case .ScreenSize 
 Case msoScreenSize800x600 
 .PixelsPerInch = 72 
 Case msoScreenSize1024x768 
 .PixelsPerInch = 96 
 Case Else 
 .PixelsPerInch = 120 
 End Select 
End With
```


## See also


[WebOptions Object](Word.WebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]