---
title: Hyperlink object (Excel)
keywords: vbaxl10.chm535072
f1_keywords:
- vbaxl10.chm535072
ms.prod: excel
api_name:
- Excel.Hyperlink
ms.assetid: 8bdd2c2f-e6eb-a2f2-78c8-b597aa80ec05
ms.date: 03/30/2019
localization_priority: Normal
---


# Hyperlink object (Excel)

Represents a hyperlink.


## Remarks

The **Hyperlink** object is a member of the **[Hyperlinks](Excel.Hyperlinks.md)** collection.


## Example

Use the **[Hyperlink](Excel.Shape.Hyperlink.md)** property of the **Shape** object to return the hyperlink for a shape (a shape can have only one hyperlink). The following example activates the hyperlink for shape one.

```vb
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```

<br/>

A range or worksheet can have more than one hyperlink. Use **[Hyperlinks](Excel.Worksheet.Hyperlinks.md)** (_index_), where _index_ is the hyperlink number, to return a single **Hyperlink** object. The following example activates hyperlink two in the range A1:B2.

```vb
Worksheets(1).Range("A1:B2").Hyperlinks(2).Follow
```


## Methods

- [AddToFavorites](Excel.Hyperlink.AddToFavorites.md)
- [CreateNewDocument](Excel.Hyperlink.CreateNewDocument.md)
- [Delete](Excel.Hyperlink.Delete.md)
- [Follow](Excel.Hyperlink.Follow.md)

## Properties

- [Address](Excel.Hyperlink.Address.md)
- [Application](Excel.Hyperlink.Application.md)
- [Creator](Excel.Hyperlink.Creator.md)
- [EmailSubject](Excel.Hyperlink.EmailSubject.md)
- [Name](Excel.Hyperlink.Name.md)
- [Parent](Excel.Hyperlink.Parent.md)
- [Range](Excel.Hyperlink.Range.md)
- [ScreenTip](Excel.Hyperlink.ScreenTip.md)
- [Shape](Excel.Hyperlink.Shape.md)
- [SubAddress](Excel.Hyperlink.SubAddress.md)
- [TextToDisplay](Excel.Hyperlink.TextToDisplay.md)
- [Type](Excel.Hyperlink.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
