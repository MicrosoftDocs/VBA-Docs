---
title: Hyperlink.EmailSubject property (PowerPoint)
keywords: vbapp10.chm526007
f1_keywords:
- vbapp10.chm526007
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.EmailSubject
ms.assetid: 2416a620-9788-5da9-3095-432cab5cdc95
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.EmailSubject property (PowerPoint)

Returns or sets the text string of the hyperlink subject line. The subject line is appended to the Internet address (URL) of the hyperlink. Read/write.


## Syntax

_expression_.**EmailSubject**

_expression_ A variable that represents an [Hyperlink](PowerPoint.Hyperlink.md) object.


## Return value

String


## Remarks

This property is commonly used with email hyperlinks. The value of this property takes precedence over any email subject specified in the  **[Address](PowerPoint.Hyperlink.Address.md)** property of the same **Hyperlink** object.


## Example

This example sets the email subject line of the first hyperlink on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Hyperlinks(1) _
    .EmailSubject = "Quote Request"
```


## See also


[Hyperlink Object](PowerPoint.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]