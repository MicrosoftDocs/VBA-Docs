---
title: Form object (Access)
keywords: vbaac10.chm13686
f1_keywords:
- vbaac10.chm13686
ms.prod: access
api_name:
- Access.Form
ms.assetid: 72ef9219-142b-b690-b696-3eba9a5d4522
ms.date: 03/08/2019
localization_priority: Priority
---


# Form object (Access)

A **Form** object refers to a particular Microsoft Access form.

## Remarks

A **Form** object is a member of the **[Forms](access.forms.md)** collection, which is a collection of all currently open forms. Within the **Forms** collection, individual forms are indexed beginning with zero. You can refer to an individual **Form** object in the **Forms** collection either by referring to the form by name, or by referring to its index within the collection. 

If you want to refer to a specific form in the **Forms** collection, it's better to refer to the form by name because a form's collection index may change. If the form name includes a space, the name must be surrounded by brackets ([ ]).

|Syntax|Example|
|:-----|:------|
|**AllForms**!_formname_|`AllForms!OrderForm`|
|**AllForms**![_form name_]|`AllForms![Order Form]`|
|**AllForms**("_formname_")|`AllForms("OrderForm")`|
|**AllForms**(_index_)|`AllForms(0)`|

Each **Form** object has a **Controls** collection, which contains all controls on the form. You can refer to a control on a form either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on the form called **OrderForm**.

```vb
 ' Implicit reference. 
Forms!OrderForm!NewData
```


```vb
' Explicit reference. 
Forms!OrderForm.Controls!NewData
```

<br/>

The next two examples show how you might refer to a control named **NewData** on a subform **ctlSubForm** contained in the form called **OrderForm**.

```vb
Forms!OrderForm.ctlSubForm.Form!Controls.NewData
```


```vb
Forms!OrderForm.ctlSubForm!NewData
```

  

## Example

The following example shows how to use **TextBox** controls to supply date criteria for a query.

```vb
Private Sub cmdSearch_Click()

   Dim db As DAO.Database
   Dim qd As QueryDef
   Dim vWhere As Variant

   Set db = CurrentDb()

   On Error Resume Next
   db.QueryDefs.Delete "Query1"
   On Error GoTo 0

   vWhere = Null

   vWhere = vWhere & " AND [PayeeID]=" + Me.cboPayeeID

   If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") <> "" Then
      vWhere = vWhere & " AND [RefundProcessed] Between #" & _
      Me.txtStartDate & "# AND #" & Me.txtEndDate & "#"
   Else
      If Nz(Me.txtEndDate, "") = "" And Nz(Me.txtStartDate, "") <> "" Then
         vWhere = vWhere & " AND [RefundProcessed]>=#" _
                  + Me.txtStartDate & "#"
      Else
         If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") = "" Then
            vWhere = vWhere & " AND [RefundProcessed] <=#" _
                     + Me.txtEndDate & "#"
      End If
     End If
   End If

   If Nz(vWhere, "") = "" Then
      MsgBox "There are no search criteria selected." & vbCrLf & vbCrLf & _
             "Search Cancelled.", vbInformation, "Search Canceled."
   Else
      Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData? & _
               " WHERE " & Mid(vWhere, 6))
      db.Close
      Set db = Nothing

      DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
   End If
End Sub
```

<br/>

The following example shows how to use the **BeforeUpdate** event of a form to require that a value be entered into one control when another control also has data.

```vb
Private Sub Form_BeforeUpdate(Cancel As Integer)
If (IsNull(Me.FieldOne)) Or (Me.FieldOne.Value =  "") Then
    ' No action required
Else
    If (IsNull(Me.FieldTwo)) or (Me.FieldTwo.Value = "") Then
        MsgBox "You must provide data for field 'FieldTwo', " & _
            "if a value is entered in FieldOne", _
            vbOKOnly, "Required Field"
        Me.FieldTwo.SetFocus
        Cancel = True
        Exit Sub
    End If
End If

End Sub
```

<br/>

The following example shows how to use the **OpenArgs** property to prevent a form from being opened from the navigation pane.

```vb
Private Sub Form_Open(Cancel As Integer)

If Me.OpenArgs() <> "Valid User" Then
    MsgBox "You are not authorized to use this form!", _
        vbExclamation + vbOKOnly, "Invalid Access"
    Cancel = True
End If
End Sub
```

<br/>

The following example shows how to use the _WhereCondition_ argument of the **OpenForm** method to filter the records displayed on a form as it is opened.

```vb
Private Sub cmdShowOrders_Click()
If Not Me.NewRecord Then
    DoCmd.OpenForm "frmOrder", _
        WhereCondition:="CustomerID=" & Me.txtCustomerID
End If
End Sub
```


## Events

- [Activate](Access.Form.Activate.md)
- [AfterDelConfirm](Access.Form.AfterDelConfirm(even).md)
- [AfterFinalRender](Access.Form.AfterFinalRender(even).md)
- [AfterInsert](Access.Form.AfterInsert(even).md)
- [AfterLayout](Access.Form.AfterLayout(even).md)
- [AfterRender](Access.Form.AfterRender(even).md)
- [AfterUpdate](Access.Form.AfterUpdate-event.md)
- [ApplyFilter](Access.Form.ApplyFilter.md)
- [BeforeDelConfirm](Access.Form.BeforeDelConfirm(even).md)
- [BeforeInsert](Access.Form.BeforeInsert(even).md)
- [BeforeQuery](Access.Form.BeforeQuery(even).md)
- [BeforeRender](Access.Form.BeforeRender(even).md)
- [BeforeScreenTip](Access.Form.BeforeScreenTip(even).md)
- [BeforeUpdate](Access.Form.BeforeUpdate-event.md)
- [Click](Access.Form.Click.md)
- [Close](Access.Form.Close.md)
- [CommandBeforeExecute](Access.Form.CommandBeforeExecute(even).md)
- [CommandChecked](Access.Form.CommandChecked(even).md)
- [CommandEnabled](Access.Form.CommandEnabled(even).md)
- [CommandExecute](Access.Form.CommandExecute(even).md)
- [Current](Access.Form.Current.md)
- [DataChange](Access.Form.DataChange(even).md)
- [DataSetChange](Access.Form.DataSetChange(even).md)
- [DblClick](Access.Form.DblClick.md)
- [Deactivate](Access.Form.Deactivate.md)
- [Delete](Access.Form.Delete.md)
- [Dirty](Access.Form.Dirty(even).md)
- [Error](Access.Form.Error.md)
- [Filter](Access.Form.Filter(even).md)
- [GotFocus](Access.Form.GotFocus.md)
- [KeyDown](Access.Form.KeyDown.md)
- [KeyPress](Access.Form.KeyPress.md)
- [KeyUp](Access.Form.KeyUp.md)
- [Load](Access.Form.Load.md)
- [LostFocus](Access.Form.LostFocus.md)
- [MouseDown](Access.Form.MouseDown.md)
- [MouseMove](Access.Form.MouseMove.md)
- [MouseUp](Access.Form.MouseUp.md)
- [MouseWheel](Access.Form.MouseWheel(even).md)
- [OnConnect](Access.Form.OnConnect(even).md)
- [OnDisconnect](Access.Form.OnDisconnect(even).md)
- [Open](Access.Form.Open.md)
- [PivotTableChange](Access.Form.PivotTableChange(even).md)
- [Query](Access.Form.Query(even).md)
- [Resize](Access.Form.Resize.md)
- [SelectionChange](Access.Form.SelectionChange(even).md)
- [Timer](Access.Form.Timer.md)
- [Undo](Access.Form.Undo(even).md)
- [Unload](Access.Form.Unload.md)
- [ViewChange](Access.Form.ViewChange(even).md)

## Methods

- [GoToPage](Access.Form.GoToPage.md)
- [Move](Access.Form.Move.md)
- [Recalc](Access.Form.Recalc.md)
- [Refresh](Access.Form.Refresh.md)
- [Repaint](Access.Form.Repaint.md)
- [Requery](Access.Form.Requery.md)
- [SetFocus](Access.Form.SetFocus.md)
- [Undo](Access.Form.Undo(method).md)

## Properties

- [ActiveControl](Access.Form.ActiveControl.md)
- [AfterDelConfirm](Access.Form.AfterDelConfirm(property).md)
- [AfterFinalRender](Access.Form.AfterFinalRender(property).md)
- [AfterInsert](Access.Form.AfterInsert(property).md)
- [AfterLayout](Access.Form.AfterLayout(property).md)
- [AfterRender](Access.Form.AfterRender(property).md)
- [AfterUpdate](Access.Form.AfterUpdate-property.md)
- [AllowAdditions](Access.Form.AllowAdditions.md)
- [AllowDatasheetView](Access.Form.AllowDatasheetView.md)
- [AllowDeletions](Access.Form.AllowDeletions.md)
- [AllowEdits](Access.Form.AllowEdits.md)
- [AllowFilters](Access.Form.AllowFilters.md)
- [AllowFormView](Access.Form.AllowFormView.md)
- [AllowLayoutView](Access.Form.AllowLayoutView.md)
- [AllowPivotChartView](Access.Form.AllowPivotChartView.md)
- [AllowPivotTableView](Access.Form.AllowPivotTableView.md)
- [Application](Access.Form.Application.md)
- [AutoCenter](Access.Form.AutoCenter.md)
- [AutoResize](Access.Form.AutoResize.md)
- [BeforeDelConfirm](Access.Form.BeforeDelConfirm(property).md)
- [BeforeInsert](Access.Form.BeforeInsert(property).md)
- [BeforeQuery](Access.Form.BeforeQuery(property).md)
- [BeforeRender](Access.Form.BeforeRender(property).md)
- [BeforeScreenTip](Access.Form.BeforeScreenTip(property).md)
- [BeforeUpdate](Access.Form.BeforeUpdate-property.md)
- [Bookmark](Access.Form.Bookmark.md)
- [BorderStyle](Access.Form.BorderStyle.md)
- [Caption](Access.Form.Caption.md)
- [ChartSpace](Access.Form.ChartSpace.md)
- [CloseButton](Access.Form.CloseButton.md)
- [CommandBeforeExecute](Access.Form.CommandBeforeExecute(property).md)
- [CommandChecked](Access.Form.CommandChecked(property).md)
- [CommandEnabled](Access.Form.CommandEnabled(property).md)
- [CommandExecute](Access.Form.CommandExecute(property).md)
- [ControlBox](Access.Form.ControlBox.md)
- [Controls](Access.Form.Controls.md)
- [Count](Access.Form.Count.md)
- [CurrentRecord](Access.Form.CurrentRecord.md)
- [CurrentSectionLeft](Access.Form.CurrentSectionLeft.md)
- [CurrentSectionTop](Access.Form.CurrentSectionTop.md)
- [CurrentView](Access.Form.CurrentView.md)
- [Cycle](Access.Form.Cycle.md)
- [DataChange](Access.Form.DataChange(property).md)
- [DataEntry](Access.Form.DataEntry.md)
- [DataSetChange](Access.Form.DataSetChange(property).md)
- [DatasheetAlternateBackColor](Access.Form.DatasheetAlternateBackColor.md)
- [DatasheetBackColor](Access.Form.DatasheetBackColor.md)
- [DatasheetBorderLineStyle](Access.Form.DatasheetBorderLineStyle.md)
- [DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)
- [DatasheetColumnHeaderUnderlineStyle](Access.Form.DatasheetColumnHeaderUnderlineStyle.md)
- [DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)
- [DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)
- [DatasheetFontName](Access.Form.DatasheetFontName.md)
- [DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)
- [DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)
- [DatasheetForeColor](Access.Form.DatasheetForeColor.md)
- [DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)
- [DatasheetGridlinesColor](Access.Form.DatasheetGridlinesColor.md)
- [DefaultControl](Access.Form.DefaultControl.md)
- [DefaultView](Access.Form.DefaultView.md)
- [Dirty](Access.Form.Dirty(property).md)
- [DisplayOnSharePointSite](Access.Form.DisplayOnSharePointSite.md)
- [DividingLines](Access.Form.DividingLines.md)
- [FastLaserPrinting](Access.Form.FastLaserPrinting.md)
- [FetchDefaults](Access.Form.FetchDefaults.md)
- [Filter](Access.Form.Filter(property).md)
- [FilterOn](Access.Form.FilterOn.md)
- [FilterOnLoad](Access.Form.FilterOnLoad.md)
- [FitToScreen](Access.Form.FitToScreen.md)
- [Form](Access.Form.Form.md)
- [FrozenColumns](Access.Form.FrozenColumns.md)
- [GridX](Access.Form.GridX.md)
- [GridY](Access.Form.GridY.md)
- [HasModule](Access.Form.HasModule.md)
- [HelpContextId](Access.Form.HelpContextId.md)
- [HelpFile](Access.Form.HelpFile.md)
- [HorizontalDatasheetGridlineStyle](Access.Form.HorizontalDatasheetGridlineStyle.md)
- [Hwnd](Access.Form.Hwnd.md)
- [InputParameters](Access.Form.InputParameters.md)
- [InsideHeight](Access.Form.InsideHeight.md)
- [InsideWidth](Access.Form.InsideWidth.md)
- [KeyPreview](Access.Form.KeyPreview.md)
- [LayoutForPrint](Access.Form.LayoutForPrint.md)
- [MaxRecButton](Access.Form.MaxRecButton.md)
- [MaxRecords](Access.Form.MaxRecords.md)
- [MenuBar](Access.Form.MenuBar.md)
- [MinMaxButtons](Access.Form.MinMaxButtons.md)
- [Modal](Access.Form.Modal.md)
- [Module](Access.Form.Module.md)
- [MouseWheel](Access.Form.MouseWheel(property).md)
- [Moveable](Access.Form.Moveable.md)
- [Name](Access.Form.Name.md)
- [NavigationButtons](Access.Form.NavigationButtons.md)
- [NavigationCaption](Access.Form.NavigationCaption.md)
- [NewRecord](Access.Form.NewRecord.md)
- [OnActivate](Access.Form.OnActivate.md)
- [OnApplyFilter](Access.Form.OnApplyFilter.md)
- [OnClick](Access.Form.OnClick.md)
- [OnClose](Access.Form.OnClose.md)
- [OnConnect](Access.Form.OnConnect(property).md)
- [OnCurrent](Access.Form.OnCurrent.md)
- [OnDblClick](Access.Form.OnDblClick.md)
- [OnDeactivate](Access.Form.OnDeactivate.md)
- [OnDelete](Access.Form.OnDelete.md)
- [OnDirty](Access.Form.OnDirty.md)
- [OnDisconnect](Access.Form.OnDisconnect(property).md)
- [OnError](Access.Form.OnError.md)
- [OnFilter](Access.Form.OnFilter.md)
- [OnGotFocus](Access.Form.OnGotFocus.md)
- [OnInsert](Access.Form.OnInsert.md)
- [OnKeyDown](Access.Form.OnKeyDown.md)
- [OnKeyPress](Access.Form.OnKeyPress.md)
- [OnKeyUp](Access.Form.OnKeyUp.md)
- [OnLoad](Access.Form.OnLoad.md)
- [OnLostFocus](Access.Form.OnLostFocus.md)
- [OnMouseDown](Access.Form.OnMouseDown.md)
- [OnMouseMove](Access.Form.OnMouseMove.md)
- [OnMouseUp](Access.Form.OnMouseUp.md)
- [OnOpen](Access.Form.OnOpen.md)
- [OnResize](Access.Form.OnResize.md)
- [OnTimer](Access.Form.OnTimer.md)
- [OnUndo](Access.Form.OnUndo.md)
- [OnUnload](Access.Form.OnUnload.md)
- [OpenArgs](Access.Form.OpenArgs.md)
- [OrderBy](Access.Form.OrderBy.md)
- [OrderByOn](Access.Form.OrderByOn.md)
- [OrderByOnLoad](Access.Form.OrderByOnLoad.md)
- [Orientation](Access.Form.Orientation.md)
- [Page](Access.Form.Page.md)
- [Pages](Access.Form.Pages.md)
- [Painting](Access.Form.Painting.md)
- [PaintPalette](Access.Form.PaintPalette.md)
- [PaletteSource](Access.Form.PaletteSource.md)
- [Parent](Access.Form.Parent.md)
- [Picture](Access.Form.Picture.md)
- [PictureAlignment](Access.Form.PictureAlignment.md)
- [PictureData](Access.Form.PictureData.md)
- [PicturePalette](Access.Form.PicturePalette.md)
- [PictureSizeMode](Access.Form.PictureSizeMode.md)
- [PictureTiling](Access.Form.PictureTiling.md)
- [PictureType](Access.Form.PictureType.md)
- [PivotTable](Access.Form.PivotTable.md)
- [PivotTableChange](Access.Form.PivotTableChange(property).md)
- [PopUp](Access.Form.PopUp.md)
- [Printer](Access.Form.Printer.md)
- [Properties](Access.Form.Properties.md)
- [PrtDevMode](Access.Form.PrtDevMode.md)
- [PrtDevNames](Access.Form.PrtDevNames.md)
- [PrtMip](Access.Form.PrtMip.md)
- [Query](Access.Form.Query(property).md)
- [RecordLocks](Access.Form.RecordLocks.md)
- [RecordSelectors](Access.Form.RecordSelectors.md)
- [Recordset](Access.Form.Recordset.md)
- [RecordsetClone](Access.Form.RecordsetClone.md)
- [RecordsetType](Access.form.recordsettype.md)
- [RecordSource](Access.Form.RecordSource.md)
- [RecordSourceQualifier](Access.Form.RecordSourceQualifier.md)
- [ResyncCommand](Access.Form.ResyncCommand.md)
- [RibbonName](Access.Form.RibbonName.md)
- [RowHeight](Access.Form.RowHeight.md)
- [ScrollBars](Access.Form.ScrollBars.md)
- [Section](Access.Form.Section.md)
- [SelectionChange](Access.Form.SelectionChange(property).md)
- [SelHeight](Access.Form.SelHeight.md)
- [SelLeft](Access.Form.SelLeft.md)
- [SelTop](Access.Form.SelTop.md)
- [SelWidth](Access.Form.SelWidth.md)
- [ServerFilter](Access.Form.ServerFilter.md)
- [ServerFilterByForm](Access.Form.ServerFilterByForm.md)
- [ShortcutMenu](Access.Form.ShortcutMenu.md)
- [ShortcutMenuBar](Access.Form.ShortcutMenuBar.md)
- [SplitFormDatasheet](Access.Form.SplitFormDatasheet.md)
- [SplitFormOrientation](Access.Form.SplitFormOrientation.md)
- [SplitFormPrinting](Access.Form.SplitFormPrinting.md)
- [SplitFormSize](Access.Form.SplitFormSize.md)
- [SplitFormSplitterBar](Access.Form.SplitFormSplitterBar.md)
- [SplitFormSplitterBarSave](Access.Form.SplitFormSplitterBarSave.md)
- [SubdatasheetExpanded](Access.Form.SubdatasheetExpanded.md)
- [SubdatasheetHeight](Access.Form.SubdatasheetHeight.md)
- [Tag](Access.Form.Tag.md)
- [TimerInterval](Access.Form.TimerInterval.md)
- [Toolbar](Access.Form.Toolbar.md)
- [UniqueTable](Access.Form.UniqueTable.md)
- [UseDefaultPrinter](Access.Form.UseDefaultPrinter.md)
- [VerticalDatasheetGridlineStyle](Access.Form.VerticalDatasheetGridlineStyle.md)
- [ViewChange](Access.Form.ViewChange(property).md)
- [ViewsAllowed](Access.Form.ViewsAllowed.md)
- [Visible](Access.Form.Visible.md)
- [Width](Access.Form.Width.md)
- [WindowHeight](Access.Form.WindowHeight.md)
- [WindowLeft](Access.Form.WindowLeft.md)
- [WindowTop](Access.Form.WindowTop.md)
- [WindowWidth](Access.Form.WindowWidth.md)



## See also

- [Access Object Model reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
