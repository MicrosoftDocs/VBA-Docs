---
title: WebCommandButton.DataRetrievalMethod property (Publisher)
keywords: vbapb10.chm3932166
f1_keywords:
- vbapb10.chm3932166
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.DataRetrievalMethod
ms.assetid: 81b89a3b-dcc5-c2b5-fbc4-6e02b587bc42
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.DataRetrievalMethod property (Publisher)

Sets or returns a **[PbSubmitDataRetrievalMethodType](Publisher.PbSubmitDataRetrievalMethodType.md)** constant that represents the way data from a web form is processed. Read/write.


## Syntax

_expression_.**DataRetrievalMethod**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

PbSubmitDataRetrievalMethodType


## Remarks

The **DataRetrievalMethod** property value can be one of the **PbSubmitDataRetrievalMethodType** constants declared in the Microsoft Publisher type library.


## Example

This example sets Microsoft Publisher to process data on the web form in the current publication by sending an email message to a specified email address.

```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]