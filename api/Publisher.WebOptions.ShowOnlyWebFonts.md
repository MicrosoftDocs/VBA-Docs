---
title: WebOptions.ShowOnlyWebFonts property (Publisher)
keywords: vbapb10.chm8257544
f1_keywords:
- vbapb10.chm8257544
ms.prod: publisher
api_name:
- Publisher.WebOptions.ShowOnlyWebFonts
ms.assetid: d18197f4-9abe-d523-77fd-f33a8ecc8076
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.ShowOnlyWebFonts property (Publisher)

Returns or sets a **Boolean** value that specifies whether only web-safe fonts and font schemes should be used when the website is viewed in a browser. If **True**, only web-safe fonts and font schemes are used. If **False**, display is not limited to web-safe fonts and font schemes. The default value is **False**. Read/write.


## Syntax

_expression_.**ShowOnlyWebFonts**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Remarks

This property applies to Latin-based fonts only.


## Example

The following example specifies that only web-safe fonts and font schemes should be used when the website is viewed in a browser.

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .ShowOnlyWebFonts = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]