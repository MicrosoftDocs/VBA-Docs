---
title: WebOptions.BrowserLevel property (Word)
keywords: vbawd10.chm165937154
f1_keywords:
- vbawd10.chm165937154
ms.prod: word
api_name:
- Word.WebOptions.BrowserLevel
ms.assetid: f753deef-cd67-918d-0fe0-af4f3d283086
ms.date: 06/08/2017
localization_priority: Normal
---


# WebOptions.BrowserLevel property (Word)

Returns or sets  **WdBrowserLevel** that represents the level of web browser at which you want to target the specified Web page. Read/write.


## Syntax

_expression_. `BrowserLevel`

_expression_ Required. A variable that represents a **[WebOptions](Word.WebOptions.md)** collection.


## Remarks

This property is ignored if the  **OptimizeForBrowser** property is set to **False**.

After you set the  **BrowserLevel** property on the **DefaultWebOptions** object, the **BrowserLevel** property of any new Web pages you create in Word will be the same as the global setting.


## See also


[WebOptions Object](Word.WebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]