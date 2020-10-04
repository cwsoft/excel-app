Attribute VB_Name = "basRibbon"
Option Explicit
Option Private Module
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basRibbon
' Callbacks and action handler for the Excel VBA Application Ribbon GUI interface (tabApp)
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
Private moAppRibbon As IRibbonUI

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' APP RIBBON CALLBACKS - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub tabApp_onLoad(Ribbon As IRibbonUI)
   ' initialize and activate tabApp ribbon at startup
   Set moAppRibbon = Ribbon
   moAppRibbon.ActivateTab "tabApp"
End Sub

Sub tabApp_Refresh()
   ' clear cached values and refresh tabApp ribbon
   On Error Resume Next
   moAppRibbon.Invalidate
End Sub

Sub tabApp_Enabled(Control As IRibbonControl, ByRef Enabled)
   ' callback to dynamically enable/disable tabApp controls
   Select Case Control.ID
      ' enable settings if logged in user is the application admin
      Case "appSettings":
         Enabled = LCase(basApp.getAppInfo("APP.Admin")) = LCase(Environ("Username"))
      
      Case Else:
         Enabled = False
   End Select
End Sub

Sub tabApp_Visible(Control As IRibbonControl, ByRef Visible)
   ' callback to dynamically show/hide tabApp controls
   Select Case Control.ID
      Case "appHelp":
         ' hide help icon if specified CHM file does not exist
         Visible = Dir(ThisWorkbook.Path & APP_CHM_PATH) <> ""
      
      Case "appContextMenu":
         ' remove the App context menu from Control and History sheet
         Visible = (ActiveSheet.Name <> APP_WKS_CONTROL And ActiveSheet.Name <> APP_WKS_HISTORY)
      
      Case Else:
         Visible = False
   End Select
End Sub

Sub tabApp_Label(ByVal Control As IRibbonControl, ByRef Label)
   ' callback to dynamically set labels of tabApp controls
   Select Case Control.ID
      Case "PutYourAppRibbonIdHere":
         Label = "your label"
      
      Case Else:
         Label = Control.ID
   End Select
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' TABAPP ACTION HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub appSave_onAction(Control As IRibbonControl)
   ' save the application
   ThisWorkbook.Save
End Sub

Sub appPrint_onAction(Control As IRibbonControl)
   ' show Excel print dialogue
   Application.Dialogs(xlDialogPrint).Show
End Sub

Sub appExit_onAction(Control As IRibbonControl)
   ' close app by seinding CTRL+w shortcut (avoids memory/application error when using ThisWorkbook.Close)
   If ThisWorkbook.Saved Then Application.DisplayAlerts = False
   Application.SendKeys "^w", True
End Sub

Sub appSettings_onAction(Control As IRibbonControl)
   ' show application settings
   Call basApp.showAppSettings
End Sub

Sub appAbout_onAction(Control As IRibbonControl)
   ' show application infos
   Call basApp.showAppInfos
End Sub

Sub appHelp_onAction(Control As IRibbonControl)
   ' show application help file
   Call basApp.showAppHelp
End Sub
