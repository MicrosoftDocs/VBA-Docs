---
title: TabStop2 object (Office)
ms.prod: office
api_name:
- Office.TabStop2
ms.assetid: fee461a9-684b-e6c2-a74a-d0aa161d0d9c
ms.date: 01/25/2019
localization_priority: Normal
---


# TabStop2 object (Office)

Represents a single tab stop. The **TabStop2** object is a member of the **[TabStops2](office.tabstops2.md)** collection.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

The following example removes the first custom tab stop from the selected paragraphs.


```vb
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub 

```


## See also

- [TabStop2 object members](overview/Library-Reference/tabstop2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]