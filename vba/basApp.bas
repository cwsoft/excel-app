Attribute VB_Name = "basApp"
Option Explicit
Option Private Module
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basApp
' Initialization, wrap-up and error handler routines for Excel based VBA Applications.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-app (https://github.com/cwsoft/excel-app)
' @requires    -
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' application wide constants
Public Const APP_DEBUG_MODE = True
Public Const APP_DISPLAY_ERRORS = True
Public Const APP_CHM_PATH = "\APP.chm"

' sheets used by this application
Public Const APP_WKS_CONTROL = "Control"
Public Const APP_WKS_HISTORY = "History"
Public Const APP_WKS_SETTINGS = "Settings"
'''

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA ROUTINES - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub Auto_Open()
   ' application initialization routines called when workbook is opened
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler

   ' register appliation shortcuts and reset Excel to defaults
   Call registerAppShortcuts(Enable:=True)
   Call setAppMode(IsInteractive:=True)
   
   ' activate App Control sheet
   With Sheets(APP_WKS_CONTROL)
      .Activate
      .Range("A1").Select
   End With
   
errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.Auto_Open")
End Sub

Sub Auto_Close()
   ' application wrap up routines called when workbook is closed
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' release appliation shortcuts and reset Excel to defaults
   Call registerAppShortcuts(Enable:=False)
   Call basApp.setAppMode(IsInteractive:=True)
   
errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.Auto_Close")
End Sub

Sub appErrorHandler(err As ErrObject, Source As String)
   ' global Application error handler
   If err.Number = 0 Or Not APP_DISPLAY_ERRORS Then Exit Sub
   
   MsgBox "An error occured in '" & Source & "'." & vbCr _
      & "Err " & CStr(err.Number) & ": " & err.Description & "." & vbCr & vbCr _
      & "Please contact the application author to fix the error." _
      , vbExclamation + vbOKOnly, "VBA Application Error"
End Sub

Sub registerAppShortcuts(Enable As Boolean)
   ' register/release application shortcuts
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   With Application
      .OnKey "^+H", IIf(Enable, "basApp.showAppHistory", "")  ' CTRL+SHIFT+H
      .OnKey "^+I", IIf(Enable, "basApp.showAppInfos", "")    ' CTRL+SHIFT+I
   End With

errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.registerAppShortcuts")
End Sub

Sub setAppMode(Optional IsInteractive, Optional StatusBarText, Optional EnableEvents)
   ' sets application properites to defined values
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   With Application
      ' set general application properties
      If Not IsEmpty(IsInteractive) Then
         .Calculation = IIf(IsInteractive, xlCalculationAutomatic, xlCalculationManual)
         .DisplayAlerts = IIf(IsInteractive, True, False)
         .ScreenUpdating = IIf(IsInteractive, True, False)
         If IsInteractive Then .EnableEvents = True
      End If
      
      ' set application status bar text
      .StatusBar = IIf(IsEmpty(StatusBarText), False, StatusBarText)
   
      ' enable/disable application events
      If Not IsEmpty(EnableEvents) Then .EnableEvents = CBool(EnableEvents)
   End With

errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.setAppMode")
End Sub

Sub showAppHelp()
   ' show application help file in CHM format (F1)
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' check if specified App help file exists
   If Dir(ThisWorkbook.Path & APP_CHM_PATH) = "" Then
      MsgBox "Application help file '" & APP_CHM_PATH & "' not found." _
         , vbExclamation + vbOKOnly, "Help file missing"
      Exit Sub
   End If
   
   On Error Resume Next
   ' display help file using Windows "hh.exe" tool
   Shell "hh " & ThisWorkbook.Path & APP_CHM_PATH, vbMaximizedFocus
   If err.Number <> 0 Then
      MsgBox "Windows tool 'hh.exe' used for displaying Windows '.chm' files is not installed on your computer." _
         , vbExclamation + vbOKOnly, "No CHM viewer installed"
   End If
   
errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.showAppHelp")
End Sub

Sub showAppHistory()
   ' show application release history tracking sheet for editing (CTRL+SHIFT+H)
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' check if logged in user is the application admin
   If LCase(basApp.getAppInfo("APP.Admin")) <> LCase(Environ("Username")) Then
      MsgBox "The application history can only by modified by the App admin '" & basApp.getAppInfo("APP.Admin") & "'." _
         , vbExclamation + vbOKOnly, "No permission to modify app history"
      Exit Sub
   End If
   
   With Sheets(APP_WKS_HISTORY)
      .Visible = xlSheetVisible
      .Activate
      .Cells(.UsedRange.Rows.Count + 2, 1).Select
   End With

errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.showAppHistory")
End Sub

Sub showAppInfos()
   ' show application infos (CTRL+SHIFT+I)
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   Load frmAppInfos
   frmAppInfos.Show

errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.showAppInfos")
End Sub

Sub showAppSettings()
   ' show application settings if logged in user the the App admin
   If APP_DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   
   ' check if logged in user is the application admin
   If LCase(basApp.getAppInfo("APP.Admin")) <> LCase(Environ("Username")) Then
      MsgBox "The application settings can only by modified by the App admin '" & basApp.getAppInfo("APP.Admin") & "'." _
         , vbExclamation + vbOKOnly, "No permission to modify app settings"
      Exit Sub
   End If
   
   With Sheets(APP_WKS_SETTINGS)
      .Visible = xlSheetVisible
      .Activate
   End With
   
errHandler:
   Call basApp.appErrorHandler(err, Source:="basApp.showAppSettings")
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA FUNCTIONS - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Function getAppInfo(Property As String)
   ' returns application info stored as workbook properties
   Dim vValue As Variant
   
   On Error Resume Next
   ' first try to extract from builin properties (File/Properties/Advanced Properties/SUMMARY)
   vValue = ""
   vValue = Trim(ThisWorkbook.BuiltinDocumentProperties(Property))
   
   ' then try to extract from custom properties (File/Properties/Advanced Properties/CUSTOM)
   If vValue = "" Then vValue = Trim(ThisWorkbook.CustomDocumentProperties(Property))
   
   err.Clear
   getAppInfo = IIf(vValue = "", "#N/A", vValue)
End Function
