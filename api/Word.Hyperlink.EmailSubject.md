---
title: Hyperlink.EmailSubject property (Word)
keywords: vbawd10.chm161285106
f1_keywords:
- vbawd10.chm161285106
ms.prod: word
api_name:
- Word.Hyperlink.EmailSubject
ms.assetid: 8b019ae2-40da-b69c-8f0b-554724a770bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.EmailSubject property (Word)

Returns or sets the text string for the specified hyperlink's subject line. Read/write  **String**.


## Syntax

_expression_.**EmailSubject**

_expression_ A variable that represents a '[Hyperlink](Word.Hyperlink.md)' object.


## Remarks

The subject line is appended to the hyperlink's Internet address, or URL. This property is commonly used with email hyperlinks. The value of this property takes precedence over any email subject specified in the  **[Address](Word.Hyperlink.Address.md)** property of the same **Hyperlink** object.


## Example

This example checks the active document for email hyperlinks; if it finds any that have a blank subject line, it adds the subject "NewProducts".


```vb
Dim hypLoop As Hyperlink 
 
For Each hypLoop In ActiveDocument.Hyperlinks 
 If hypLoop.Address Like "mailto*" And _ 
 hypLoop.Address = hypLoop.EmailSubject Then 
 hypLoop.EmailSubject = "NewProducts" 
 End If 
Next hypLoop
```


## See also


[Hyperlink Object](Word.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]