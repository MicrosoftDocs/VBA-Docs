---
title: COMAddIns object (Office)
keywords: vbaof11.chm220000
f1_keywords:
- vbaof11.chm220000
ms.prod: office
api_name:
- Office.COMAddIns
ms.assetid: f6efa1cc-8d30-27d5-8b07-7ddad22f16ef
ms.date: 01/03/2019
localization_priority: Normal
---


# COMAddIns object (Office)

A collection of **[COMAddIn](Office.COMAddIn.md)** objects that provide information about a COM add-in registered in the Windows registry.

## Example

Use the **COMAddIns** property of the **Application** object to return the **COMAddIns** collection for a Microsoft Office host application. This collection contains all of the COM add-ins that are available to a given Office host application, and the **Count** property of the **COMAddins** collection returns the number of available COM add-ins, as in the following example.

```vb
MsgBox Application.COMAddIns.Count
```

<br/>

Use the **Update** method of the **COMAddins** collection to refresh the list of COM add-ins from the Windows registry, as in the following example.

```vb
Application.COMAddIns.Update
```

<br/>

Use **COMAddIns.Item(index)**, where _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text and ProgID (`"msodraa9.ShapeSelect"`) in a message box.

```vb
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```

## See also

- [COMAddIns object members](overview/Library-Reference/comaddins-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]