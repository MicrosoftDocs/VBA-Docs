---
title: RemoveItem method (Microsoft Forms)
keywords: fm20.chm5224968
f1_keywords:
- fm20.chm5224968
ms.prod: office
api_name:
- Office.RemoveItem
ms.assetid: b895775c-7b77-6f2b-b368-998d7114aa7a
ms.date: 11/15/2018
localization_priority: Normal
---


# RemoveItem method (Microsoft Forms)

Removes a row from the list in a list box or combo box.

## Syntax

_Boolean_ = _object_. **RemoveItem**_index_

The **RemoveItem** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. Specifies the row to delete. The number of the first row is 0; the number of the second row is 1, and so on.|

This method will not remove a row from the list if the **[ListBox](listbox-control.md)** is data-[bound](../../Glossary/glossary-vba.md#bound) (that is, when the **RowSource** property specifies a [data source](../../Glossary/glossary-vba.md#data-source) for the **[ListBox](listbox-control.md)**).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]