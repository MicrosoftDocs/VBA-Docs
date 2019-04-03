---
title: ListFormat object (Word)
keywords: vbawd10.chm2496
f1_keywords:
- vbawd10.chm2496
ms.prod: word
api_name:
- Word.ListFormat
ms.assetid: 74773fd6-b713-34d4-b7be-f543c983008d
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat object (Word)

Represents the list formatting attributes that can be applied to the paragraphs in a range.


## Remarks

Use the  **ListFormat** property to return the **ListFormat** object for a range. The following example applies the default bulleted list format to the selection.


```vb
Selection.Range.ListFormat.ApplyBulletDefault
```

An easy way to apply list formatting is to use the  **ApplyBulletDefault**, **ApplyNumberDefault**, and **ApplyOutlineNumberDefault** methods, which correspond, respectively, to the first list format (excluding **None**) on each tab in the  **Bullets and Numbering** dialog box.

To apply a format other than the default format, use the  **ApplyListTemplate** method, which allows you to specify the list format (list template) you want to apply.

Use the  **List** or **ListTemplate** property to return the list or list template from the first paragraph in the specified range.

Use the  **ListFormat** property with a **Range** object to access the list formatting properties and methods available for the specified range. The following example applies the default bulleted list format to the second paragraph in the active document.




```vb
ActiveDocument.Paragraphs(2).Range.ListFormat.ApplyBulletDefault
```

However, if there is already a list defined in your document, you can access a  **List** object by using the **Lists** property. The following example changes the format of the list created in the preceding example to the first number format on the **Numbered** tab in the **Bullets and Numbering** dialog box.




```vb
ActiveDocument.Lists(1).ApplyListTemplate _ 
 ListTemplate:=ListGalleries(2).ListTemplates(1)
```


## Methods



|Name|
|:-----|
|[ApplyBulletDefault](Word.ListFormat.ApplyBulletDefault.md)|
|[ApplyListTemplate](Word.ListFormat.ApplyListTemplate.md)|
|[ApplyListTemplateWithLevel](Word.ListFormat.ApplyListTemplateWithLevel.md)|
|[ApplyNumberDefault](Word.ListFormat.ApplyNumberDefault.md)|
|[ApplyOutlineNumberDefault](Word.ListFormat.ApplyOutlineNumberDefault.md)|
|[CanContinuePreviousList](Word.ListFormat.CanContinuePreviousList.md)|
|[ConvertNumbersToText](Word.ListFormat.ConvertNumbersToText.md)|
|[CountNumberedItems](Word.ListFormat.CountNumberedItems.md)|
|[ListIndent](Word.ListFormat.ListIndent.md)|
|[ListOutdent](Word.ListFormat.ListOutdent.md)|
|[RemoveNumbers](Word.ListFormat.RemoveNumbers.md)|

## Properties



|Name|
|:-----|
|[Application](Word.ListFormat.Application.md)|
|[Creator](Word.ListFormat.Creator.md)|
|[List](Word.ListFormat.List.md)|
|[ListLevelNumber](Word.ListFormat.ListLevelNumber.md)|
|[ListPictureBullet](Word.ListFormat.ListPictureBullet.md)|
|[ListString](Word.ListFormat.ListString.md)|
|[ListTemplate](Word.ListFormat.ListTemplate.md)|
|[ListType](Word.ListFormat.ListType.md)|
|[ListValue](Word.ListFormat.ListValue.md)|
|[Parent](Word.ListFormat.Parent.md)|
|[SingleList](Word.ListFormat.SingleList.md)|
|[SingleListTemplate](Word.ListFormat.SingleListTemplate.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
