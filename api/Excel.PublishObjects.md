---
title: PublishObjects object (Excel)
keywords: vbaxl10.chm649072
f1_keywords:
- vbaxl10.chm649072
ms.prod: excel
api_name:
- Excel.PublishObjects
ms.assetid: 33ad393e-5ab6-2531-5e5b-42930fc596c0
ms.date: 03/30/2019
localization_priority: Normal
---


# PublishObjects object (Excel)

A collection of all **[PublishObject](Excel.PublishObject.md)** objects in the workbook.


## Remarks

Each **PublishObject** object represents an item in a workbook that has been saved to a webpage and can be refreshed according to values specified by the properties and methods of the object.


## Example

Use the **[PublishObjects](Excel.Workbook.PublishObjects.md)** property of the **Workbook** object to return the **PublishObjects** collection. The following example saves all static **PublishObject** objects in the active workbook to the webpage.

```vb
Set objPObjs = ActiveWorkbook.PublishObjects 
For Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```

<br/>

Use **PublishObjects** (_index_), where _index_ is the index number of the specified item in the workbook, to return a single **PublishObject** object. The following example sets the location where the first item in workbook three is saved.

```vb
Workbooks(3).PublishObjects(1).FileName = _ 
 "\\myserver\public\finacct\statemnt.htm"
```

## Methods

- [Add](Excel.PublishObjects.Add.md)
- [Delete](Excel.PublishObjects.Delete.md)
- [Publish](Excel.PublishObjects.Publish.md)

## Properties

- [Application](Excel.PublishObjects.Application.md)
- [Count](Excel.PublishObjects.Count.md)
- [Creator](Excel.PublishObjects.Creator.md)
- [Item](Excel.PublishObjects.Item.md)
- [Parent](Excel.PublishObjects.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]