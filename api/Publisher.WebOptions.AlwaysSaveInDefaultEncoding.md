---
title: WebOptions.AlwaysSaveInDefaultEncoding property (Publisher)
keywords: vbapb10.chm8257539
f1_keywords:
- vbapb10.chm8257539
ms.prod: publisher
api_name:
- Publisher.WebOptions.AlwaysSaveInDefaultEncoding
ms.assetid: e37ff08f-5c09-0a71-27e1-e2a332147087
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.AlwaysSaveInDefaultEncoding property (Publisher)

Returns or sets a **Boolean** value that specifies whether webpages within a web publication should always be saved by using default encoding. If **True**, webpages within a publication will always be saved by using the default encoding of the client computer. If **False**, webpages will not be saved by using default encoding. The default value is **False**. Read/write.


## Syntax

_expression_.**AlwaysSaveInDefaultEncoding**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Remarks

If the **AlwaysSaveInDefaultEncoding** property is set to **True** on a given **WebOptions** object, any subsequent attempts to set the **[Encoding](Publisher.WebOptions.Encoding.md)** property on that object will be ignored.


## Example

The following example tests whether the web publication is currently set to be saved by using default encoding. If so, the **AlwaysSaveInDefaultEncoding** property is set to **False**, and the **Encoding** property is used to set the encoding to Unicode (UTF-8).

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .AlwaysSaveInDefaultEncoding = True Then 
 .AlwaysSaveInDefaultEncoding = False 
 .Encoding = msoEncodingUTF8 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]