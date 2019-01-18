---
title: TextColumn2 object (Office)
ms.prod: office
api_name:
- Office.TextColumn2
ms.assetid: 631387c1-2b7a-6c98-d05f-c054434c8b9d
ms.date: 06/08/2017
localization_priority: Normal
---


# TextColumn2 object (Office)

Represents a single text column. The  **TextColumn2** object is a member of the **TextColumns2** collection.


## Remarks

Use  **TextColumns2(Index)**, where _Index_ is the index number, to return a single **TextColumn2** object. The index number represents the position of the column in the **TextColumns2** collection (counting from left to right).


## Example

Use the  **Add** method to add a column to the collection of columns. By default, there's one text column in the **TextColumns2** collection. The following example adds a 2.5-inch-widecolumn to the active Microsoft Word document.


```vb
ActiveDocument.PageSetup.TextColumns2.Add _ 
 Width:=InchesToPoints(2.5), _ 
 Spacing:=InchesToPoints(0.5), EvenlySpaced:=False 

```


## Properties



|Name|
|:-----|
|[Application](Office.TextColumn2.Application.md)|
|[Creator](Office.TextColumn2.Creator.md)|
|[Number](Office.TextColumn2.Number.md)|
|[Spacing](Office.TextColumn2.Spacing.md)|
|[TextDirection](Office.TextColumn2.TextDirection.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]