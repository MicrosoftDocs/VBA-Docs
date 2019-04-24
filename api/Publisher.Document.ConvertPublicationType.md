---
title: Document.ConvertPublicationType method (Publisher)
keywords: vbapb10.chm196737
f1_keywords:
- vbapb10.chm196737
ms.prod: publisher
api_name:
- Publisher.Document.ConvertPublicationType
ms.assetid: e4bfe349-a22f-6017-ac9d-49f67e1f6dd2
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ConvertPublicationType method (Publisher)

Converts the specified publication to the specified publication type.


## Syntax

_expression_.**ConvertPublicationType**(**_Value_**)

 _expression_ A variable that represents a  **Document** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Value|Required| **PbPublicationType**|The type of publication to which you want the publication converted.|

## Remarks

When a publication is converted, any settings that apply to its previous type remain, but are ignored. For example, converting a print publication to a web publication results in any advanced print settings being ignored. If the publication is converted back to a print publication, the settings take effect again.

Use the  **[PublicationType](Publisher.Document.PublicationType.md)** property of the **[Document](Publisher.Document.md)** object to determine the publication type of a publication.

The Value parameter can be one of the following  **PbPublicationType** constants declared in the Microsoft Publisher type library.



| **pbTypePrint**|
| **pbTypeWeb**|

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