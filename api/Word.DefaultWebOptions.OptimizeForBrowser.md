---
title: DefaultWebOptions.OptimizeForBrowser property (Word)
keywords: vbawd10.chm165871617
f1_keywords:
- vbawd10.chm165871617
ms.prod: word
api_name:
- Word.DefaultWebOptions.OptimizeForBrowser
ms.assetid: c85aced0-0f4d-8237-e9c9-15fc65e0fd2b
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.OptimizeForBrowser property (Word)

 **True** if Microsoft Word optimizes new Web pages created in Word for the Web browser specified by the **[BrowserLevel](Word.DefaultWebOptions.BrowserLevel.md)** property. Read/write **Boolean**.


## Syntax

 _expression_. `OptimizeForBrowser`

 _expression_ Required. A variable that represents a '[DefaultWebOptions](Word.DefaultWebOptions.md)' collection.


## Example

This example sets Word to optimize new Web pages for Microsoft Internet Explorer 5 and creates a Web page based on this setting.


```vb
With Application.DefaultWebOptions 
 .BrowserLevel _ 
 = wdBrowserLevelMicrosoftInternetExplorer5 
 .OptimizeForBrowser = True 
End With 
Documents.Add DocumentType:=wdNewWebPage
```


## See also


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]