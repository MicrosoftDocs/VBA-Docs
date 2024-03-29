---
title: WebOptions.Encoding property (Excel)
keywords: vbaxl10.chm662082
f1_keywords:
- vbaxl10.chm662082
api_name:
- Excel.WebOptions.Encoding
ms.assetid: 99395ad8-4503-eac2-b194-6a4706e5264d
ms.date: 05/18/2019
ms.localizationpriority: medium
---


# WebOptions.Encoding property (Excel)

Returns or sets the document encoding (code page or character set) to be used by the web browser when you view the saved document. The default is the system code page. Read/write **[MsoEncoding](Office.MsoEncoding.md)**.


## Syntax

_expression_.**Encoding**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

You cannot use any of the constants that have the suffix **AutoDetect**. These constants are used by the **[ReloadAs](Excel.Workbook.ReloadAs.md)** method.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]