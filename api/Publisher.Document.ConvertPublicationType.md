---
title: Document.ConvertPublicationType method (Publisher)
keywords: vbapb10.chm196737
f1_keywords:
- vbapb10.chm196737
ms.prod: publisher
api_name:
- Publisher.Document.ConvertPublicationType
ms.assetid: e4bfe349-a22f-6017-ac9d-49f67e1f6dd2
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ConvertPublicationType method (Publisher)

Converts the specified publication to the specified publication type.


## Syntax

_expression_.**ConvertPublicationType** (_Value_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **[PbPublicationType](publisher.pbpublicationtype.md)**|The type of publication to which you want the publication converted. Can be one of the **PbPublicationType** constants declared in the Microsoft Publisher type library.|

## Remarks

When a publication is converted, any settings that apply to its previous type remain, but are ignored. For example, converting a print publication to a web publication results in any advanced print settings being ignored. If the publication is converted back to a print publication, the settings take effect again.

Use the **[PublicationType](Publisher.Document.PublicationType.md)** property to determine the publication type of a publication.

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