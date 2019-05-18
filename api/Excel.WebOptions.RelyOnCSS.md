---
title: WebOptions.RelyOnCSS property (Excel)
keywords: vbaxl10.chm662073
f1_keywords:
- vbaxl10.chm662073
ms.prod: excel
api_name:
- Excel.WebOptions.RelyOnCSS
ms.assetid: 7921e4d8-f27f-4e87-e64e-ae0f8c5869c3
ms.date: 05/18/2019
localization_priority: Normal
---


# WebOptions.RelyOnCSS property (Excel)

**True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your webpage, depending on the value of the **[OrganizeInFolder](Excel.WebOptions.OrganizeInFolder.md)** property. **False** if HTML `<FONT>` tags and cascading style sheets are used. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**RelyOnCSS**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

You should set this property to **True** if your web browser supports cascading style sheets because this gives you more precise layout and formatting control on your webpage and makes it look more like your document (as it appears in Microsoft Excel).


## Example

This example enables the use of cascading style sheets as the global default for the application.

```vb
ThisWorkbook.WebOptions.RelyOnCSS = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]