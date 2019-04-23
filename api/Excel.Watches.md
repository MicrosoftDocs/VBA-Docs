---
title: Watches object (Excel)
keywords: vbaxl10.chm687072
f1_keywords:
- vbaxl10.chm687072
ms.prod: excel
api_name:
- Excel.Watches
ms.assetid: de403bcc-b927-90f6-75d7-9c936c7f58f7
ms.date: 04/03/2019
localization_priority: Normal
---


# Watches object (Excel)

A collection of all the **[Watch](Excel.Watches.md)** objects in a specified application.


## Example

Use the **[Watches](Excel.Application.Watches.md)** property of the **Application** object to return a **Watches** collection.

In the following example, Microsoft Excel creates a new **Watch** object by using the **Add** method. This example creates a summation formula in cell A3, and then adds this cell to the watch facility.

```vb
Sub AddWatch() 
 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.Add Source:=ActiveCell 
 End With 
 
End Sub
```

<br/>

You can specify to remove individual cells from the watch facility by using the **Delete** method of the **Watches** collection. This example deletes cell A3 on worksheet 1 of book 1 from the Watch window. This example assumes that you have added the cell A3 on sheet 1 of book 1 (by using the previous example to add a **Watch** object).

```vb
Sub DeleteAWatch() 
 
 Application.Watches(Workbooks("Book1").Sheets("Sheet1").Range("A3")).Delete 
 
End Sub
```

<br/>

You can also specify to remove all cells from the Watch window by using the **Delete** method of the **Watches** collection. This example deletes all cells from the Watch window.

```vb
Sub DeleteAllWatches() 
 
 Application.Watches.Delete 
 
End Sub
```

## Methods

- [Add](Excel.Watches.Add.md)
- [Delete](Excel.Watches.Delete.md)

## Properties

- [Application](Excel.Watches.Application.md)
- [Count](Excel.Watches.Count.md)
- [Creator](Excel.Watches.Creator.md)
- [Item](Excel.Watches.Item.md)
- [Parent](Excel.Watches.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]