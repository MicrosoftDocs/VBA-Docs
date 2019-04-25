---
title: Hyperlink.EmailSubject property (Excel)
keywords: vbaxl10.chm536083
f1_keywords:
- vbaxl10.chm536083
ms.prod: excel
api_name:
- Excel.Hyperlink.EmailSubject
ms.assetid: 3fe6d6a1-8184-8ef5-eb6e-b96ce9732dbd
ms.date: 04/26/2019
localization_priority: Normal
---


# Hyperlink.EmailSubject property (Excel)

Returns or sets the text string of the specified hyperlink's email subject line. The subject line is appended to the hyperlink's address. Read/write **String**.


## Syntax

_expression_.**EmailSubject**

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


## Remarks

This property is usually used with email hyperlinks.

The value of this property takes precedence over any email subject line that you have specified by using the **[Address](Excel.Hyperlink.Address.md)** property of the **Hyperlink** object.


## Example

This example sets the email subject line for the first hyperlink in the first worksheet.

```vb
Worksheets(1).Hyperlinks(1).EmailSubject = "Quote Request"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]