---
title: DefaultWebOptions.PixelsPerInch property (Word)
keywords: vbawd10.chm165871628
f1_keywords:
- vbawd10.chm165871628
ms.prod: word
api_name:
- Word.DefaultWebOptions.PixelsPerInch
ms.assetid: baae93ab-1e1e-79ae-1717-3671367a34cc
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.PixelsPerInch property (Word)

Returns or sets the density (pixels per inch) of graphics images and table cells on a webpage. Read/write  **Long**.


## Syntax

_expression_.**PixelsPerInch**

_expression_ Required. A variable that represents a **[DefaultWebOptions](Word.DefaultWebOptions.md)** collection.


## Remarks

The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96.

This property determines the size of the images and cells on the specified Web page relative to the size of text whenever you view the saved document in a web browser. The physical dimensions of the resulting image or cell are the result of the original dimensions (in inches) multiplied by the number of pixels per inch.

Use the **ScreenSize** property to set the optimum screen size for the targeted web browsers.


## Example

This example sets the pixel density depending on the target screen size of the web browser.


```vb
With Application.DefaultWebOptions 
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


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]