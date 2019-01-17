---
title: UsedObjects object (Excel)
keywords: vbaxl10.chm677072
f1_keywords:
- vbaxl10.chm677072
ms.prod: excel
api_name:
- Excel.UsedObjects
ms.assetid: b94ad3d1-411f-acf6-19bb-8e6c4a484748
ms.date: 06/08/2017
localization_priority: Normal
---


# UsedObjects object (Excel)

Represents objects that have been allocated in a workbook.


## Example

Use the  **[UsedObjects](Excel.Application.UsedObjects.md)** property of the **[Application](Excel.Application(object).md)** object to return a **UsedObjects** collection.

Once a  **UsedObjects** collection is returned, you can determine the quantity of used objects in a Microsoft Excel application using the **[Count](Excel.UsedObjects.Count.md)** property.



In this example, Microsoft Excel determines the quantity of objects that have been allocated and notifies the user. This example assumes a recalculation was performed in the application and was interrupted before finishing.






```vb
Sub CountUsedObjects() 
 
 MsgBox "The number of used objects in this application is: " & _ 
 Application.UsedObjects.Count 
 
End Sub
```


## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]