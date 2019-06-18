---
title: WebOptions.RelyOnVML property (Publisher)
keywords: vbapb10.chm8257543
f1_keywords:
- vbapb10.chm8257543
ms.prod: publisher
api_name:
- Publisher.WebOptions.RelyOnVML
ms.assetid: 8cd29d64-48a6-d33e-cb9d-6b1ea094069a
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.RelyOnVML property (Publisher)

Returns or sets a **Boolean** value that specifies whether image files are generated from drawing objects when a web publication is saved. If **True**, image files are not generated. If **False**, images are generated. The default value is **False**. Read/write.


## Syntax

_expression_.**RelyOnVML**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Remarks

File sizes can be reduced by not generating images for drawing objects. Note that a web browser must support Vector Markup Language (VML) to render drawing objects. Microsoft Internet Explorer versions 5.0 and later support VML, so the **RelyOnVML** property could be set to **True** if targeting those browsers. For browsers that do not support VML, a drawing object will not appear when a web publication is saved with this property enabled.

If unsure about which browsers will be used to view the website, this property should be set to **False**.


## Example

The following example assumes that users have Microsoft Internet Explorer version 5.0, and therefore specifies that images should not be generated from drawing objects when the web publication is saved.

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .RelyOnVML = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]