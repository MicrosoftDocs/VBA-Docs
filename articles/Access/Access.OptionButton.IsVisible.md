---
title: OptionButton.IsVisible Property (Access)
keywords: vbaac10.chm10606
f1_keywords:
- vbaac10.chm10606
ms.prod: access
api_name:
- Access.OptionButton.IsVisible
ms.assetid: e9fdcd98-275a-7e54-bee5-74d97a6de086
ms.date: 06/08/2017
---


# OptionButton.IsVisible Property (Access)

You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.


## Syntax

 _expression_. **IsVisible**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks


 **Note**  You can set the  **IsVisible** property only in the **Print** event of a report section that contains the control.

You can use the  **IsVisible** property together with the **HideDuplicates** property to determine when a control on a report is visible and show or hide other controls as a result. For example, you could hide a line control when a text box control is hidden because it contains duplicate values.


## Example

The following example uses the  **IsVisible** property of a text box to control the display of a line control on a report. T



Paste the following code into the Declarations section of the report module, and then view the report to see the line formatting controlled by the  **IsVisible** property:




```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 If Me!CategoryID.IsVisible Then 
 Me!Line0.Visible = True 
 Else 
 Me!Line0.Visible = False 
 End If 
End Sub
```


## See also


#### Concepts


[OptionButton Object](Access.OptionButton.md)

