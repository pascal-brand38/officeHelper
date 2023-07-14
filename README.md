# compta.xla

xla is the add-in complement macros of excel.
This file contains macros that are used by compta.xls, a proprietary excel file.

In order to use it, please add the following in compta.xls. Replace %enterprise% with your proprietary name.

## In VBAProject (compta.xls) / Microsoft Excel Objects / ThisWorkbook

```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Run "comptaHelperComputeGainsParMois"
End Sub

Private Sub Workbook_Open()
  Workbooks.Open Filename:="C:\msys64\home\virgi\dev\pascal-brand38\officeHelper\comptaHelper.xla"
  Run "comptaHelperInit", "%enterprise%", "Agenda %enterprise%\%enterprise%-agenda.ods", "%enterprise%-contrat.pdf"
  
  Run "comptaHelperComputeGainsParMois"
  Run "comptaHelperCreateCommandBar"
End Sub
```

## In VBAProject (compta.xls) / Microsoft Excel Objects / Feuil1 (Gains par mois)

```
Private Sub Worksheet_Activate()
  Run "comptaHelperComputeGainsParMois"
End Sub
```