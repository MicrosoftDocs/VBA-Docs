---
title: Application.DefaultWebOptions property (Excel)
keywords: vbaxl10.chm133247
f1_keywords:
- vbaxl10.chm133247
ms.prod: excel
api_name:
- Excel.Application.DefaultWebOptions
ms.assetid: 51524888-0812-85ee-c8f9-e14d9b558f57
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DefaultWebOptions property (Excel)

Returns the  **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object that contains global application-level attributes used by Microsoft Excel whenever you save a document as a Web page or open a Web page. Read-only.


## Syntax

_expression_. `DefaultWebOptions`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example checks to see whether the default setting for document encoding is Western, and then it sets the string  `strDocEncoding` accordingly.


```vb
If Application.DefaultWebOptions.Encoding = msoEncodingWestern Then 
 strDocEncoding = "Western" 
Else 
 strDocEncoding = "Other" 
End If
```


## See also


[Application Object](Excel.Application(object).md)

