---
title: IRibbonUI Object (Office)
keywords: vbaof11.chm320000
f1_keywords:
- vbaof11.chm320000
ms.prod: office
api_name:
- Office.IRibbonUI
ms.assetid: d323aa21-de74-e821-c914-db71ef3b9c5e
ms.date: 08/29/2018
---


# IRibbonUI Object (Office)

The object that is returned by the **onLoad** procedure specified on the **customUI** tag. The object contains methods for invalidating control properties and for refreshing the user interface.

The iRibbonUI does not generate Events in its interaction with the user. Instead, ribbon elements perform *callbacks* to your code, and the linkage between ribbon elements and your code is defined in the XML that describes your ribbon additions. For information about the callback functions available for each UI element, see [Customizing the 2007 Office Fluent Ribbon for Developers](https://docs.microsoft.com/previous-versions/office/developer/office-2007/aa722523(v=office.12))  and search for "How can I determine the correct signatures for each callback procedure?"



## Remarks

You can customize the Ribbon user interface (UI) by using callback procedures in COM add-ins. When the host application starts, the **onLoad** callback procedure is called. The callback procedure then returns a **IRibbonUI** object pointing to the user interface (UI). YOu can use that object to invoke the **Invalidate**, **InvalidateControl**, and **Refresh** methods.


## Example

In the following example, starting the host application triggers the **onLoad** event procedure that then calls a procedure which creates a **IRibbonUI** object representing the Ribbon UI. Next, a callback procedure is defined that invalidates all of the cached controls and then refreshes the UI.


```xml
<customUI … OnLoad="MyAddInInitialize" …>
```


```vb
Dim MyRibbon As IRibbonUI 
 
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.Invalidate() ' Invalidates the caches of all of this add-in's controls 
End Sub
```


## See also


[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
[RibbonXML Callbacks](https://docs.microsoft.com/previous-versions/office/developer/office-2007/aa722523(v=office.12))
[IRibbonUI Object Members](./overview/Library-Reference/iribbonui-members-office.md)

