---
title: Range.ShowAll property (Word)
keywords: vbawd10.chm157155736
f1_keywords:
- vbawd10.chm157155736
ms.prod: word
api_name:
- Word.Range.ShowAll
ms.assetid: 751077ec-5ea4-c60a-ac92-d8a5a3c13620
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.ShowAll property (Word)

 **True** if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write **Boolean**.


## Syntax

_expression_. `ShowAll`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

This property only affects the specified range when  **Show Markup** is set to **Show Revisions in Balloons**. When  **Range.ShowAll** is set to **False**,  **[Range.Text](Word.Range.Text.md)** provides all of the text in the range except deleted text. If you set **Range.ShowAll** to **True**, then  **[Range.Text](Word.Range.Text.md)** provides all of the text in the range including the text that was deleted.


## Example

The following example displays all the text in the specified range, excluding deleted text.


> [!NOTE] 
> This example assumes that the active document has change tracking enabled, that  **Show Markup** is set to **Show Revisions in Balloons**, and that some text has been deleted from the document.


```vb
Sub HideDeletedText()
Dim r As Range

Set r = ActiveDocument.Range
r.ShowAll = False
Debug.Print r.Text

End Sub
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]