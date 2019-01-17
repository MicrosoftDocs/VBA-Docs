---
title: Using events with the QueryTable object
keywords: vbaxl10.chm5255039
f1_keywords:
- vbaxl10.chm5255039
ms.prod: excel
ms.assetid: 9f58dcba-1832-2aa5-be03-0ce85d0c5cd1
ms.date: 11/13/2018
localization_priority: Normal
---


# Using events with the QueryTable object

Before you can use events with the **QueryTable** object, you must first create a class module and declare a **QueryTable** object with events. For example, assume that you have created a class module and named it `ClsModQT`. This module contains the following code:


```vb
Public WithEvents qtQueryTable As QueryTable
```

After you have declared the new object by using events, it appears in the **Object** list box in the class module.

Before the procedures will run, however, you must connect the declared object in the class module to the specified **QueryTable** object. You can do this by entering the following code in the class module:

```vb
Sub InitQueryEvent(QT as Object) 
 Set qtQueryTable = QT 
End Sub
```

After you run this initialization procedure, the object you declared in the class module points to the specified **QueryTable** object. You can initialize the event in a module by calling the event. In this example, the first query table on the active worksheet is connected to the **qtQueryTable** object.

```vb
Dim clsQueryTable as New ClsModQT 
 
Sub RunInitQTEvent 
 clsQueryTable.InitQueryEvent _ 
 QT:=ActiveSheet.QueryTables(1) 
End Sub
```

You can write other event procedures in the object's class. When you click the new object in the **Object** box, the valid events for that object are displayed in the **Procedure** list box.

## See also

- [Excel functions (by category)](https://support.office.com/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]