---
title: DefaultWebOptions.BrowserLevel property (Word)
keywords: vbawd10.chm165871618
f1_keywords:
- vbawd10.chm165871618
ms.prod: word
api_name:
- Word.DefaultWebOptions.BrowserLevel
ms.assetid: 15817831-8921-df0b-43fc-43bad18116d6
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.BrowserLevel property (Word)

Returns or sets a  **WdBrowserLevel** constant that represents the level of the web browser for which you want to target new Web pages created in Microsoft Word. Read/write.


## Syntax

_expression_. `BrowserLevel`

_expression_ Required. A variable that represents a **[DefaultWebOptions](Word.DefaultWebOptions.md)** collection.


## Remarks

After you set the  **BrowserLevel** property on the **DefaultWebOptions** object, the **BrowserLevel** property of any new Web pages you create in Word will be the same as the global setting.


## Example

This example sets Word to optimize new Web pages for Microsoft Internet Explorer 5 and creates a webpage based on this setting.


```vb
With Application.DefaultWebOptions 
 .BrowserLevel = wdBrowserLevelMicrosoftInternetExplorer5 
 .OptimizeForBrowser = True 
End With 
Documents.Add DocumentType:=wdNewWebPage
```


## See also


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]