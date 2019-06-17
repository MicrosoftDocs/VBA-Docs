---
title: WebOptions.EnableIncrementalUpload property (Publisher)
keywords: vbapb10.chm8257541
f1_keywords:
- vbapb10.chm8257541
ms.prod: publisher
api_name:
- Publisher.WebOptions.EnableIncrementalUpload
ms.assetid: 42d5e22e-7e39-848e-a7e7-5d2019f7e71c
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.EnableIncrementalUpload property (Publisher)

Returns or sets a **Boolean** value that specifies whether changes made to a web publication can be uploaded to a web server independent of the entire publication. If **True**, only changes made to a publication are uploaded to the web server when published. If **False**, the entire publication is uploaded to the web server. The default value is **True**. Read/write.


## Syntax

_expression_.**EnableIncrementalUpload**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Remarks

The **EnableIncrementalUpload** property applies only to web publications that have already been published to a web server. If a web publication has not already been published to a web server, the entire publication will be published to the server during the initial publishing process, regardless of whether the **EnableIncrementalUpload** property is set to **True**.

If a web publication has already been published to a web server and the **EnableIncrementalUpload** property is then set to **True**, only changes made to the web publication, and not the entire publication, after this point will be published to the server.


## Example

The following example tests whether the web publication is set to upload only changes made to the publication. If not, the **EnableIncrementalUpload** property is set to **True** to specify that only changes to the publication be uploaded to the web server.

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .EnableIncrementalUpload = False Then 
 .EnableIncrementalUpload = True 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]