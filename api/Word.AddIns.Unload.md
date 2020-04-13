---
title: AddIns.Unload method (Word)
keywords: vbawd10.chm159318019
f1_keywords:
- vbawd10.chm159318019
ms.prod: word
api_name:
- Word.AddIns.Unload
ms.assetid: de0e4683-2630-0d2b-03d7-7710be1a6740
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns.Unload method (Word)

Unloads all loaded add-ins and, depending on the value of the  _RemoveFromList_ argument, removes them from the **AddIns** collection.


## Syntax

_expression_. `Unload`( `_RemoveFromList_` )

 _expression_ An expression that returns an '[AddIns](Word.addins.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RemoveFromList_|Required| **Boolean**| **True** to remove the unloaded add-ins from the **AddIns** collection (the names are removed from the **Templates and Add-ins** dialog box). **False** to leave the unloaded add-ins in the collection. If the **Autoload** property for an unloaded add-in returns **True**, **Unload** cannot remove that add-in from the **AddIns** collection, regardless of the value of RemoveFromList.|

## Remarks

To unload a single template or WLL, set the **[Installed](Word.AddIn.Installed.md)** property of the **AddIn** object to **False**. To remove a single template or WLL from the **AddIns** collection, apply the **[Delete](Word.AddIn.Delete.md)** method to the **AddIn** object.


## Example

This example unloads all the add-ins listed in the **Templates and Add-ins** dialog box. The add-in names remain in the **AddIns** collection.


```vb
If AddIns.Count > 0 Then AddIns.UnLoad RemoveFromList:=False
```


## See also


[AddIns Collection Object](Word.addins.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]