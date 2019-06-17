---
title: WebOptions.Encoding property (Publisher)
keywords: vbapb10.chm8257540
f1_keywords:
- vbapb10.chm8257540
ms.prod: publisher
api_name:
- Publisher.WebOptions.Encoding
ms.assetid: 0aad6082-0ee4-3be0-14a0-73e219f254a0
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.Encoding property (Publisher)

Returns an **[MsoEncoding](Office.MsoEncoding.md)** constant that specifies the encoding of the web publication. Read/write.


## Syntax

_expression_.**Encoding**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

MsoEncoding


## Remarks

If the **[AlwaysSaveInDefaultEncoding](Publisher.WebOptions.AlwaysSaveInDefaultEncoding.md)** property is set to **True** on a given **WebOptions** object, any subsequent attempts to set the **Encoding** property on that object will be ignored.

The **Encoding** property value can be one of the **MsoEncoding** constants declared in the Microsoft Office type library.

Attempting to set the **Encoding** property to an **MsoEncoding** constant that is not available on the client computer results in a run-time error.

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