---
title: TabStops2 object (Office)
ms.prod: office
api_name:
- Office.TabStops2
ms.assetid: 1d1d8054-19eb-cd65-f37d-36e93e7fc347
ms.date: 06/08/2017
---


# TabStops2 object (Office)

The collection of  **TabStop2** objects.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

 The following example removes the first custom tab stop from the first paragraph in the active Microsoft Publisher publication.


```vb
Sub ClearTabStop() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
        .ParagraphFormat.Tabs(1).Clear 
End Sub 

```


## Methods



|Name|
|:-----|
|[Add](Office.TabStops2.Add.md)|
|[Item](Office.TabStops2.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Office.TabStops2.Application.md)|
|[Count](Office.TabStops2.Count.md)|
|[Creator](Office.TabStops2.Creator.md)|
|[DefaultSpacing](Office.TabStops2.DefaultSpacing.md)|
|[Parent](Office.TabStops2.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
