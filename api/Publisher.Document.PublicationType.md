---
title: Document.PublicationType property (Publisher)
keywords: vbapb10.chm196736
f1_keywords:
- vbapb10.chm196736
ms.prod: publisher
api_name:
- Publisher.Document.PublicationType
ms.assetid: 264c2769-2452-0009-4853-84a6a426db38
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.PublicationType property (Publisher)

Returns a **[PbPublicationType](publisher.pbpublicationtype.md)** constant that represents the type of the specified publication. Read-only.


## Syntax

_expression_.**PublicationType**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

PbPublicationType


## Example

The following example determines if the active publication is a print publication. If it is, the publication is converted to a web publication.

```vb
Sub ChangePublicationType() 
 With ActiveDocument 
 If .PublicationType = pbTypePrint Then 
 .ConvertPublicationType (pbTypeWeb) 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]