---
title: TabStop2 Object (Office)
ms.prod: office
api_name:
- Office.TabStop2
ms.assetid: fee461a9-684b-e6c2-a74a-d0aa161d0d9c
ms.date: 06/08/2017
---


# TabStop2 Object (Office)

Represents a single tab stop. The  **TabStop2** object is a member of the **TabStops2** collection.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

The following example removes the first custom tab stop from the selected paragraphs.


```vb
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Clear](Office.TabStop2.Clear.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.TabStop2.Application.md)|
|[Creator](Office.TabStop2.Creator.md)|
|[Parent](Office.TabStop2.Parent.md)|
|[Position](Office.TabStop2.Position.md)|
|[Type](Office.TabStop2.Type.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
