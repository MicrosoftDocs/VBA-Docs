---
title: Application.CommandBars property (Visio)
keywords: vis_sdr.chm10050540
f1_keywords:
- vis_sdr.chm10050540
ms.prod: visio
api_name:
- Visio.Application.CommandBars
ms.assetid: 3829b033-aed4-a132-ff44-96d419dd09cd
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.CommandBars property (Visio)

Returns a reference to the **[CommandBars](office.commandbars.md)** collection that represents the command bars in the container application. Read-only.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Syntax

_expression_.**CommandBars**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

CommandBars


## Remarks

Beginning with Microsoft Visio 2002, a program can manipulate menus and toolbars in the Visio user interface by manipulating the **CommandBars** collection returned by the **CommandBars** property. The **CommandBars** collection has an interface identical to the **CommandBars** collection exposed by the suite of Microsoft Office applications such as Microsoft Word and Microsoft Excel.

Alternatively, since Visio version 4.0, Visio has exposed application and document properties that return a **[UIObject](visio.uiobject.md)** object that provides similar functionality to **CommandBars**. Consequently, programs can use either the **CommandBars** collection or **UIObject** objects to manipulate the Visio menus and toolbars.

To get information about the object returned by the **CommandBars** property:

1. On the **Developer** tab, choose **Visual Basic**.
    
2. On the **View** menu, choose **Object Browser**.
    
3. In the **Project/Library** list, choose **Office**.
    
4. If you do not see the Office type library in the **Project/Library** list, on the **Tools** menu, choose **References**, select the **Microsoft Office 14.0 Object Library** check box, and then choose **OK**.
    
5. Under **Classes**, examine the class named **CommandBars**.
    

Each **CommandBarControl** object in a **CommandBars** collection has an **OnAction** property, and each **CommandBar** object in a **CommandBars** collection has a **Context** property. The values of these properties are determined by the container application. In Microsoft Visio:

- The **OnAction** property is a **String** value that is interpreted either as a COM add-in, as a Microsoft Visual Basic for Applications (VBA) macro, as VBA code, or as a Visio add-on name.
    
- The **Context** property determines in which menu context a command bar appears. 

  - The menu context number is a **String** value (for example **visUIObjSetDrawing** or "2"), which is followed by an asterisk if the command bar is visible by default (for example, **visUIObjSetShapeSheet** `& "*" or "4*"`). 
  - Valid menu contexts are **visUIObjSetDrawing** (2), **visUIObjSetStencil** (3), **visUIObjSetShapeSheet** (4), **visUIObjSetIcon** (5), or **visUIObjSetPrintPreview** (7). Attempting to set the **Context** property to any other value will fail.

## Example

This macro shows how to use the **CommandBars** property to list the command bars.

```vb
 
Public Sub CommandBars_Example() 
 
 Dim vsoCommandBars As CommandBars 
 Dim vsoCommandBar As CommandBar 
 
 'Get the set of CommandBars 
 'for the application. 
 Set vsoCommandBars = Application.CommandBars 
 
 'List each CommandBar in the Immediate window. 
 For Each vsoCommandBar In vsoCommandBars 
 Debug.Print vsoCommandBar.Name 
 Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]