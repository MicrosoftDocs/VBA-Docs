---
title: DocumentInspectors object (Office)
keywords: vbaof11.chm278000
f1_keywords:
- vbaof11.chm278000
ms.prod: office
api_name:
- Office.DocumentInspectors
ms.assetid: 8366d7cd-e016-bb99-d27f-749ca10352f1
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentInspectors object (Office)

Represents a collection of **[DocumentInspector](office.documentinspector.md)** objects.


## Remarks

The **DocumentInspectors** collection is part of the **Document** object in Microsoft Word, the **Workbook** object in Excel, and the **Presentation** object in PowerPoint. A **DocumentInspectors** collection contains multiple **DocumentInspector** objects, one for some built-in options and each installed custom Document Inspector module. 


## Example

The following example calls the **Fix** method of a Document Inspector module and displays the status of the action and the specific items that are removed.


```vb
Public Sub FixDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim results As String 
 ActiveDocument.DocumentInspectors(3).Fix docStatus, results 
 
 MsgBox docStatus 
 MsgBox("The following items were removed " & results) 
 
End Sub 

```


## See also

- [DocumentInspectors object members](overview/library-reference/documentinspectors-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]