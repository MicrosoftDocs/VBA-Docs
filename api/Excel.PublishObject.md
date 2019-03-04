---
title: PublishObject object (Excel)
keywords: vbaxl10.chm651072
f1_keywords:
- vbaxl10.chm651072
ms.prod: excel
api_name:
- Excel.PublishObject
ms.assetid: da719d86-b65b-3bbd-c0fc-8b3113777540
ms.date: 06/08/2017
localization_priority: Normal
---


# PublishObject object (Excel)

Represents an item in a workbook that has been saved to a web page and can be refreshed according to values specified by the properties and methods of the  **PublishObject** object.


## Remarks

 The **PublishObject** object is a member of the **[PublishObjects](Excel.PublishObjects.md)** collection.


## Example

Use  **[PublishObjects](Excel.Workbook.PublishObjects.md)** ( _index_ ), where _index_ is the index number of the specified item in the workbook, to return a single **PublishObject** object. The following example sets the location where the first item in workbook three is saved.


```vb
Workbooks(3).PublishObjects(1).FileName = _ 
 "\\myserver\public\finacct\statemnt.htm"
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]