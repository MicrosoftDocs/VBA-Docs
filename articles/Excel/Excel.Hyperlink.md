---
title: Hyperlink Object (Excel)
keywords: vbaxl10.chm535072
f1_keywords:
- vbaxl10.chm535072
ms.prod: excel
api_name:
- Excel.Hyperlink
ms.assetid: 8bdd2c2f-e6eb-a2f2-78c8-b597aa80ec05
ms.date: 06/08/2017
---


# Hyperlink Object (Excel)

Represents a hyperlink.


## Remarks

 The **Hyperlink** object is a member of the **[Hyperlinks](Excel.Hyperlinks.md)** collection.


## Example

Use the  **[Hyperlink](Excel.Shape.Hyperlink.md)** property to return the hyperlink for a shape (a shape can have only one hyperlink). The following example activates the hyperlink for shape one.


```
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```

A range or worksheet can have more than one hyperlink. Use  **[Hyperlinks](Excel.Worksheet.Hyperlinks.md)** ( _index_ ), where _index_ is the hyperlink number, to return a single **Hyperlink** object. The folllowing example activates hyperlink two in the range A1:B2.




```
Worksheets(1).Range("A1:B2").Hyperlinks(2).Follow
```


## Methods



|**Name**|
|:-----|
|[AddToFavorites](Excel.Hyperlink.AddToFavorites.md)|
|[CreateNewDocument](Excel.Hyperlink.CreateNewDocument.md)|
|[Delete](Excel.Hyperlink.Delete.md)|
|[Follow](Excel.Hyperlink.Follow.md)|

## Properties



|**Name**|
|:-----|
|[Address](Excel.Hyperlink.Address.md)|
|[Application](Excel.Hyperlink.Application.md)|
|[Creator](Excel.Hyperlink.Creator.md)|
|[EmailSubject](Excel.Hyperlink.EmailSubject.md)|
|[Name](Excel.Hyperlink.Name.md)|
|[Parent](Excel.Hyperlink.Parent.md)|
|[Range](Excel.Hyperlink.Range.md)|
|[ScreenTip](Excel.Hyperlink.ScreenTip.md)|
|[Shape](Excel.Hyperlink.Shape.md)|
|[SubAddress](Excel.Hyperlink.SubAddress.md)|
|[TextToDisplay](Excel.Hyperlink.TextToDisplay.md)|
|[Type](Excel.Hyperlink.Type.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
