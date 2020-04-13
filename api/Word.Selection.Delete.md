---
title: Selection.Delete method (Word)
keywords: vbawd10.chm158662783
f1_keywords:
- vbawd10.chm158662783
ms.prod: word
api_name:
- Word.Selection.Delete
ms.assetid: 35bfdf19-62d3-5593-0b2f-dd6b642b4cc3
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Delete method (Word)

Deletes the specified number of characters or words.


## Syntax

_expression_.**Delete**( `_Unit_` , `_Count_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The unit by which the collapsed selection is to be deleted. Can be one of the **WdUnits** constants.|
| _Count_|Optional| **Variant**|The number of units to be deleted. To delete units after the selection, collapse the selection and use a positive number. To delete units before the selection, collapse the selection and use a negative number.|

## Return value

Long


## Remarks

This method returns a  **Long** value that indicates the number of items deleted, or it returns 0 (zero) if the deletion was unsuccessful.


## Example

This example selects and deletes the contents of the active document.


```vb
Sub DeleteSelection() 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Are you sure you want to " & _ 
 "delete the contents of the document?", vbYesNo) 
 
 If intResponse = vbYes Then 
 ActiveDocument.Content.Select 
 Selection.Delete 
 End If 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
