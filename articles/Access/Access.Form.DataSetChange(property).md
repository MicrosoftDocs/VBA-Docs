---
title: Form.DataSetChange Property (Access)
keywords: vbaac10.chm13546,vbaac10.chm5111
f1_keywords:
- vbaac10.chm13546,vbaac10.chm5111
ms.prod: access
api_name:
- Access.Form.DataSetChange
ms.assetid: 29f7f9a8-4dbd-9f69-7f4c-7f93add9f1b6
ms.date: 06/08/2017
---


# Form.DataSetChange Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[DataSetChange](Access.Form.DataSetChange(even).md)** event occurs. Read/write.


## Syntax

 _expression_. **DataSetChange**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **DataSetChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).DataSetChange = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](Access.Form.md)

