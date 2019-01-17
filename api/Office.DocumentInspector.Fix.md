---
title: DocumentInspector.Fix method (Office)
keywords: vbaof11.chm279004
f1_keywords:
- vbaof11.chm279004
ms.prod: office
api_name:
- Office.DocumentInspector.Fix
ms.assetid: b05326b0-779c-97f5-d3fd-705f82a141ef
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentInspector.Fix method (Office)

Performs an action on specific information items or document properties depending on the Document Inspector module specified.


## Syntax

_expression_.**Fix**(_Status_, _Results_)

_expression_ An expression that returns a **[DocumentInspector](Office.DocumentInspector.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Status_|Required|**[MsoDocInspectorStatus](office.msodocinspectorstatus.md)**|An enumeration representing the status of the document. _Status_ is an output parameter, which means that its value is returned when the method has completed its purpose.|
| _Results_|Required|**String**|Contains the results of the action. _Results_ is an output parameter.|

## Remarks

There are two Document Inspector modules that are included with Microsoft Office. These are the **Comments/Revisions** module and the **Document Properties** method. These are the first two options that show up in the **Document Inspector** dialog box, but are not available in the **DocumentInspectors** collection.


## Example

The following example demonstrates implementing the **Fix** method of the **DocumentInspector** object. You specify which Document Inspector module to execute with the index value specified in the **DocumentInspectors** collection. Before executing this method, you would likely run the **Inspect** method to determine if there are any hidden worksheets in the workbook.


```vb
Public Sub DI_FixDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim result As String 
ActiveDocument.DocumentInspectors(3).Fix docStatus, result 
 
MsgBox ("The Fix method returned the following status " &amp; docStatus &amp; _ 
" with this result " &amp; result) 
End Sub
```


## See also

- [DocumentInspector object members](overview/library-reference/documentinspector-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]