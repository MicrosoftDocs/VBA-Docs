---
title: CustomTaskPane.VisibleStateChange event (Office)
keywords: vbaof11.chm302001
f1_keywords:
- vbaof11.chm302001
ms.prod: office
api_name:
- Office.CustomTaskPane.VisibleStateChange
ms.assetid: 6faccef7-f35f-d0c8-383f-54493e4b4c8b
ms.date: 01/04/2019
localization_priority: Normal
---


# CustomTaskPane.VisibleStateChange event (Office)

Occurs when the user changes the visibility of the custom task pane.


## Syntax

_expression_.**VisibleStateChange** (_CustomTaskPaneInst_)

_expression_ An expression that returns a **[CustomTaskPane](Office.CustomTaskPane.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CustomTaskPaneInst_|Required|**CustomTaskPane**|The active task pane.|

## Example

The following example, written in C#, creates a custom task pane and adds an ActiveX button control created in another project. A **VisibleStateChange** event of type **_CustomTaskPaneEvents_VisibleStateChangeEventHandler** is defined in the procedure. When the event is triggered, the event handler displays a message box depending on whether the task pane is currently visible or hidden.


```cs
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
 
 CTP.VisibleStateChange += new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(CTP_VisibleStateChange); 
} 
 
private void CTP_VisibleStateChange(object sender, string visiblestateArgs) 
{ 
 if (CTP.Visible) 
 { 
 Console.WriteLine("The custom task pane is now visible"); 
 } 
 else 
 { 
 Console.WriteLine("The custom task pane has been hidden"); 
 } 
} 

```


> [!NOTE] 
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also

- [CustomTaskPane object members](overview/library-reference/customtaskpane-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]