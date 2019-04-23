---
title: TabStops2 object (Office)
ms.prod: office
api_name:
- Office.TabStops2
ms.assetid: 1d1d8054-19eb-cd65-f37d-36e93e7fc347
ms.date: 01/25/2019
localization_priority: Normal
---


# TabStops2 object (Office)

The collection of **[TabStop2](Office.TabStop2.md)** objects.


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


## See also

- [TabStops2 object members](overview/Library-Reference/tabstops2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]