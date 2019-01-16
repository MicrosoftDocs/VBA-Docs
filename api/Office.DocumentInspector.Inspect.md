---
title: DocumentInspector.Inspect method (Office)
keywords: vbaof11.chm279003
f1_keywords:
- vbaof11.chm279003
ms.prod: office
api_name:
- Office.DocumentInspector.Inspect
ms.assetid: 5973fa7d-7218-74e3-b67c-c03fbaf4b930
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentInspector.Inspect method (Office)

Inspects a document for specific information or document properties.


## Syntax

_expression_.**Inspect**(_Status_, _Results_)

_expression_ An expression that returns a **[DocumentInspector](Office.DocumentInspector.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Status_|Required|**[MsoDocInspectorStatus](office.msodocinspectorstatus.md)**|An enumeration representing the status of the document. _Status_ is an output parameter, which means that its value is returned when the method has completed its purpose.|
| _Results_|Required|**String**|Contains a list of the information items or document properties found in the document.|


## Example

The following example inspects a document by using the **Inspect** method of the **DocumentInspector** object, and then displays the status and results of the inspection.


```vb
Public Sub DI_InspectDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim result As String 
ActiveDocument.DocumentInspectors(1).Inspect docStatus, results 
 
MsgBox ("The inspection returned the following status " &amp; docStatus &amp; _ 
" with this result " &amp; result) 
End Sub
```


## See also

- [DocumentInspector object members](overview/library-reference/documentinspector-members-office.md)

