---
title: Hyperlink.ScreenTip property (Excel)
keywords: vbaxl10.chm536084
f1_keywords:
- vbaxl10.chm536084
ms.prod: excel
api_name:
- Excel.Hyperlink.ScreenTip
ms.assetid: 472aeaca-90f4-3b27-6927-a51d708e61c2
ms.date: 04/26/2019
localization_priority: Normal
---


# Hyperlink.ScreenTip property (Excel)

Returns or sets the ScreenTip text for the specified hyperlink. Read/write **String**.


## Syntax

_expression_.**ScreenTip**

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


## Remarks

After the document has been saved to a webpage, the ScreenTip text may appear (for example) when the mouse pointer is positioned over the hyperlink while the document is being viewed in a web browser. Some web browsers may not support ScreenTips.


## Example

This example sets the screen tip for the first hyperlink on the active worksheet.

```vb
ActiveSheet.Hyperlinks(1).ScreenTip = "Return to the home page"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]