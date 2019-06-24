---
title: Application.DefaultWebOptions method (Word)
keywords: vbawd10.chm158335381
f1_keywords:
- vbawd10.chm158335381
ms.prod: word
api_name:
- Word.Application.DefaultWebOptions
ms.assetid: ee683d3c-b331-cccd-27ec-b3258b42961e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DefaultWebOptions method (Word)

Returns the  **[DefaultWebOptions](Word.DefaultWebOptions.md)** object that contains global application-level attributes used by Microsoft Word whenever you save a document as a webpage or open a webpage.


## Syntax

_expression_. `DefaultWebOptions`

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Return value

DefaultWebOptions


## Example

This example checks to see whether the default setting for document encoding is Western, and then it sets the string strDocEncoding accordingly.


```vb
Dim strDocEncoding As String 
 
If Application.DefaultWebOptions.Encoding _ 
 = msoEncodingWestern Then 
 strDocEncoding = "Western" 
Else 
 strDocEncoding = "Other" 
End If
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]