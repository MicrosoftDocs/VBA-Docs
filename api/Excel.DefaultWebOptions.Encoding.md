---
title: DefaultWebOptions.Encoding property (Excel)
keywords: vbaxl10.chm660086
f1_keywords:
- vbaxl10.chm660086
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.Encoding
ms.assetid: 53164ab3-b0f5-ed8e-76f8-840cbd8e23bc
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.Encoding property (Excel)

Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. The default is the system code page. Read/write **[MsoEncoding](Office.MsoEncoding.md)**.


## Syntax

_expression_.**Encoding**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

You cannot use any of the constants that have the suffix **AutoDetect**. These constants are used by the **[ReloadAs](Excel.Workbook.ReloadAs.md)** method.


## Example

This example checks to see whether the default document encoding is Western, and then it sets the string `strDocEncoding` accordingly.

```vb
If Application.DefaultWebOptions.Encoding = msoEncodingWestern Then 
    strDocEncoding = "Western" 
Else 
    strDocEncoding = "Other" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]