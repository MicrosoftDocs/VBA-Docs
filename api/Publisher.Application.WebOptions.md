---
title: Application.WebOptions Property (Publisher)
keywords: vbapb10.chm131176
f1_keywords:
- vbapb10.chm131176
ms.prod: publisher
api_name:
- Publisher.Application.WebOptions
ms.assetid: 2e0c3435-a55a-4903-a0f8-9c347dec03b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WebOptions Property (Publisher)

Returns a  **[WebOptions](Publisher.WebOptions.md)** object, which represents the properties of Web publications. Read-only.


## Syntax

 _expression_. **WebOptions**

 _expression_ A variable that represents a  **Application** object.


## Return value

WebOptions


## Example

The following example specifies that Web publications should not always be saved in default encoding, and that the encoding should be Unicode (UTF-8).


```vb
With Application.WebOptions 
 .AlwaysSaveInDefaultEncoding = False 
 .Encoding = msoEncodingUTF8 
End With
```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]