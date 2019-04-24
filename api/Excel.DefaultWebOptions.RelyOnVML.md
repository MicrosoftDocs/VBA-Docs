---
title: DefaultWebOptions.RelyOnVML property (Excel)
keywords: vbaxl10.chm660081
f1_keywords:
- vbaxl10.chm660081
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.RelyOnVML
ms.assetid: 414b81f1-2549-f9b3-a5a0-b36dcbf8a6e4
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.RelyOnVML property (Excel)

**True** if image files are not generated from drawing objects when you save a document as a webpage. **False** if images are generated. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**RelyOnVML**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

You can reduce file sizes by not generating images for drawing objects, if your web browser supports Vector Markup Language (VML). For example, Microsoft Internet Explorer 5 supports this feature, and you should set the **RelyOnVML** property to **True** if you are targeting this browser. For browsers that do not support VML, the image will not appear when you view a webpage saved with this property enabled.

For example, you should not generate images if your webpage uses image files that you have generated earlier, and if the location where you save the document is different from the final location of the page on the web server.


## Example

This example specifies that images are generated when saving the worksheet to a webpage.

```vb
Workbooks(1).DefaultWebOptions.RelyOnVML = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]