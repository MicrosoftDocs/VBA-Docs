---
title: ToggleButton.HideDuplicates Property (Access)
keywords: vbaac10.chm11717
f1_keywords:
- vbaac10.chm11717
ms.prod: access
api_name:
- Access.ToggleButton.HideDuplicates
ms.assetid: 3bcd4798-81fa-0cfb-4dd4-1ed9150dbb3a
ms.date: 06/08/2017
---


# ToggleButton.HideDuplicates Property (Access)

You can use the  **HideDuplicates** property to hide a control on a report when its value is the same as in the preceding record. Read/write **Boolean**.


## Syntax

 _expression_. **HideDuplicates**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **HideDuplicates** property applies only to controls (check box, combo box, list box, option button, option group, text box, toggle button) on a report.

The  **HideDuplicates** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|If the value of a control or the data it contains is the same as in the preceding record, the control is hidden.|
|**False**|(Default) The control is visible regardless of the value in the preceding record.|
The  **DefaultValue** property doesn't apply to check box, option button, or toggle buttoncontrols when they are in an option group. It does however apply to the option group itself.

You can set the  **HideDuplicates** property only in report Design view.

You can use the  **HideDuplicates** property to create a grouped report by using only the detail section rather than a group header and the detail section.


## Example

The following example returns the  **HideDuplicates** property setting for the CategoryName text box and assigns the value to the `intCurVal` variable.


```vb
Dim intCurVal As Integer 
intCurVal = Me!CategoryName.HideDuplicates
```


## See also


#### Concepts


[ToggleButton Object](Access.ToggleButton.md)

