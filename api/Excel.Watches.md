---
title: Watches object (Excel)
keywords: vbaxl10.chm687072
f1_keywords:
- vbaxl10.chm687072
ms.prod: excel
api_name:
- Excel.Watches
ms.assetid: de403bcc-b927-90f6-75d7-9c936c7f58f7
ms.date: 06/08/2017
localization_priority: Normal
---


# Watches object (Excel)

A collection of all the  **[Watch](Excel.Watches.md)** objects in a specified application.


## Example

Use the  **[Watches](Excel.Application.Watches.md)** property of the **[Application](Excel.Application(object).md)** object to return a **Watches** collection.



In the following example, Microsoft Excel creates a new  **Watch** object using the **[Add](Excel.Watches.Add.md)** method. This example creates a summation formula in cell A3, and then adds this cell to the watch facility.




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

You can specify to remove individual cells from the watch facility by using the  **[Delete](Excel.Watches.Delete.md)** method of the **Watches** collection. This example deletes cell A3 on worksheet 1 of book 1 from the Watch Window. This example assumes you have added the cell A3 on sheet 1 of book 1 (using the previous example to add a **Watch** object).




```vb
Sub DeleteAWatch() 
 
 Application.Watches(Workbooks("Book1").Sheets("Sheet1").Range("A3")).Delete 
 
End Sub
```

You can also specify to remove all cells from the Watch Window, by using the  **Delete** method of the **Watches** collection. This example deletes all cells from the Watch Window.




```vb
Sub DeleteAllWatches() 
 
 Application.Watches.Delete 
 
End Sub
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

