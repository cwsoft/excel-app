VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAppInfos 
   Caption         =   "Application Infos"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7995
   OleObjectBlob   =   "frmAppInfos.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAppInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Form: frmAppInfos
' Code to display the release history of the excel based application.
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

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' FORM EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub UserForm_Initialize()
   ' initialize form before beeing displayed
   Dim rng As Range
   Dim sAppHistory As String
   
   On Error Resume Next
   With Me
      ' extract application infos from Excel settings
      .lblAppName = basApp.getAppInfo("Title")
      .lblAppAuthor = basApp.getAppInfo("Author")
      .lblAppVersion = basApp.getAppInfo("App.Version")
      .lblAppDescription = basApp.getAppInfo("Comments")
      
      ' extract application history
      sAppHistory = ""
      For Each rng In Sheets(APP_WKS_HISTORY).Range("A3:A" & Sheets(APP_WKS_HISTORY).UsedRange.Rows.Count)
         sAppHistory = sAppHistory & CStr(rng.Value) & vbCr
      Next
      .txtAppReleaseHistory.Text = sAppHistory
   End With
End Sub
