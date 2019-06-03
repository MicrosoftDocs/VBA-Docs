---
title: DocumentInspector object (Office)
keywords: vbaof11.chm279000
f1_keywords:
- vbaof11.chm279000
ms.prod: office
api_name:
- Office.DocumentInspector
ms.assetid: 75dcf0ca-5afa-996b-e8d2-13d71ac0f6f8
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentInspector object (Office)

Represents a Document Inspector module in a **[DocumentInspectors](Office.DocumentInspectors.md)** collection.


## Remarks

The **DocumentInspector** object provides access to the **[Inspect](Office.DocumentInspector.Inspect.md)** and **[Fix](Office.DocumentInspector.Fix.md)** methods. 

The **DocumentInspector** object represents custom Document Inspector modules and some "built-in" options. To see the list of built-in options, open the **Document Inspector** dialog box from the **File** menu. The first two options (**Comments, Revisions, Versions, and Annotations** and **Document Properties and Personal Information**) are not listed in the **DocumentInspectors** collection; instead, their functionality is available through the **RemoveDocumentInformation** method. 

The remaining options in the dialog box and any installed custom modules developed by third-party sources are available from the **DocumentInspectors** collection by specifying an index value.


## Example

The following example inspects a document by using a Document Inspector module, and displays the status and results of the inspection.


```vb
Public Sub DI_InspectDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim result As String 
ActiveDocument.DocumentInspectors(3).Inspect docStatus, results 
 
MsgBox ("The inspection returned the following status " & docStatus & _ 
" with this result " & result) 
End Sub
```


## See also

- [DocumentInspector object members](overview/library-reference/documentinspector-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]