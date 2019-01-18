---
title: DefaultWebOptions.ScreenSize property (Word)
keywords: vbawd10.chm165871627
f1_keywords:
- vbawd10.chm165871627
ms.prod: word
api_name:
- Word.DefaultWebOptions.ScreenSize
ms.assetid: 21f1019f-6658-0da9-519e-adefc8356607
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.ScreenSize property (Word)

Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Read/write  **MsoScreenSize**.


## Syntax

 _expression_. `ScreenSize`

 _expression_ Required. A variable that represents a '[DefaultWebOptions](Word.DefaultWebOptions.md)' collection.


## Example

This example sets the target screen size at 800x600 pixels.


```vb
Application.DefaultWebOptions.ScreenSize = _ 
 msoScreenSize800x600
```


## See also


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

