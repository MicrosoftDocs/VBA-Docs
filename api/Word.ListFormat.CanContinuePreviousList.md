---
title: ListFormat.CanContinuePreviousList method (Word)
keywords: vbawd10.chm163578040
f1_keywords:
- vbawd10.chm163578040
ms.prod: word
api_name:
- Word.ListFormat.CanContinuePreviousList
ms.assetid: 5c9a91e4-999e-d976-126d-673831f2ecaf
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.CanContinuePreviousList method (Word)

Returns a  **WdContinue** constant (**wdContinueDisabled**, **wdResetList**, or **wdContinueList**) that indicates whether the formatting from the previous list can be continued.


## Syntax

_expression_. `CanContinuePreviousList`( `_ListTemplate_` )

_expression_ Required. A variable that represents a '[ListFormat](Word.ListFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ListTemplate_|Required| **[ListTemplate](Word.ListTemplate.md)**|A list template that's been applied to previous paragraphs in the document.|

## Remarks

This method returns the state of the  **Continue previous list** and **Restart numbering** options in the **Bullets and Numbering** dialog box for a specified list format. To change the settings of these options, set the ContinuePreviousList argument of the **ApplyListTemplate** method.




## Example

This example checks to see whether numbering from a previous list is disabled. If it isn't disabled, the current list template is applied with numbering set to continue from the previous list. The selection must be within the second list, or this example creates an error.


```vb
Dim lfTemp As ListFormat 
Dim intContinue As Integer 
 
Set lfTemp = Selection.Range.ListFormat 
 
intContinue = lfTemp.CanContinuePreviousList( _ 
 ListTemplate:=lfTemp.ListTemplate) 
If intContinue <> wdContinueDisabled Then 
 lfTemp.ApplyListTemplate _ 
 ListTemplate:=lfTemp.ListTemplate, _ 
 ContinuePreviousList:=True 
End If
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]