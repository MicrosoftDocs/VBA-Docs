---
title: List Object (Word)
keywords: vbawd10.chm2450
f1_keywords:
- vbawd10.chm2450
ms.prod: word
api_name:
- Word.List
ms.assetid: 2c3dae28-447a-af48-2966-e19ae75ab6c2
ms.date: 06/08/2017
---


# List Object (Word)

Represents a single list format that's been applied to specified paragraphs in a document. The  **List** object is a member of the **Lists** collection.


## Remarks

Use  **Lists** (Index), where Index is the index number, to return a single **List** object. The following example returns the number of items in list one in the active document.


```
mycount = ActiveDocument.Lists(1).CountNumberedItems
```

To return all the paragraphs that have list formatting, use the  **ListParagraphs** property. To return them as a range, use the **Range** property.

To apply a different list format to an existing list, use the  **ApplyListTemplate** method with the **List** object. To add a new list to a document, use the **ApplyListTemplate** method with the **[ListFormat](Word.ListFormat.md)** object for a specified range.

Use the  **CanContinuePreviousList** method to determine whether you can continue the list formatting from a list that was previously applied to the document.

Use the  **CountNumberedItems** method to return the number of items in a numbered or bulleted list, including LISTNUM fields.

To determine whether a list contains more than one list template, use the  **SingleListTemplate** property.

You can manipulate the individual  **List** objects within a document, but for more precise control you should work with the **ListFormat** object.


 **Note**  Picture-bulleted lists are not included in the  **[Lists](Word.lists.md)** collection and cannot be manipulated using the **List** object.


## Methods



|**Name**|
|:-----|
|[ApplyListTemplate](Word.List.ApplyListTemplate.md)|
|[ApplyListTemplateWithLevel](Word.List.ApplyListTemplateWithLevel.md)|
|[CanContinuePreviousList](Word.List.CanContinuePreviousList.md)|
|[ConvertNumbersToText](Word.List.ConvertNumbersToText.md)|
|[CountNumberedItems](Word.List.CountNumberedItems.md)|
|[RemoveNumbers](Word.List.RemoveNumbers.md)|

## Properties



|**Name**|
|:-----|
|[Application](Word.List.Application.md)|
|[Creator](Word.List.Creator.md)|
|[ListParagraphs](Word.List.ListParagraphs.md)|
|[Parent](Word.List.Parent.md)|
|[Range](Word.List.Range.md)|
|[SingleListTemplate](Word.List.SingleListTemplate.md)|
|[StyleName](Word.List.StyleName.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
