---
title: CustomTaskPane.DockPositionStateChange event (Office)
keywords: vbaof11.chm302002
f1_keywords:
- vbaof11.chm302002
ms.prod: office
api_name:
- Office.CustomTaskPane.DockPositionStateChange
ms.assetid: fd22407b-4926-2de5-ec1d-aad1a13fe269
ms.date: 01/04/2019
localization_priority: Normal
---


# CustomTaskPane.DockPositionStateChange event (Office)

Occurs when the user changes the docking position of the active custom task pane.


## Syntax

_expression_.**DockPositionStateChange** (_CustomTaskPaneInst_)

_expression_ An expression that returns a **[CustomTaskPane](Office.CustomTaskPane.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CustomTaskPaneInst_|Required|**Object**|The active custom task pane.|

## Example

The following example, written in C#, creates a custom task pane and adds a Microsoft ActiveX button control that was created in another project. A **DockPositionStateChange** event of type **_CustomTaskPaneEvents_DockPositionStateChangeEventHandler** is then defined. When the event is triggered, a message box is displayed telling the user that the docked task pane has been moved.

```cs
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
 
 CTP.DockPositionStateChange += new _CustomTaskPaneEvents_DockPositionStateChangeEventHandler(CTP_DockPositionStateChange); 
 
} 
 
private void CTP_DockPositionStateChange(object sender, string dockpositionArgs) 
{ 
 Console.WriteLine("The custom task pane was moved"); 
}
```


> [!NOTE] 
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 

## See also

- [CustomTaskPane object members](overview/library-reference/customtaskpane-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]