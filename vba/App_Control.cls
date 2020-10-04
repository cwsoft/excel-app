VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sheet: App_Control
' This sheet serves as main entry point of the Excel based application.
' Following the steps ensures that the Excel based App works as intented.
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
' SHEET EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdResetApplication_Click()
   ' reset application
   MsgBox "Add your App function here.", vbInformation + vbOKOnly, "To be implemented"
End Sub

Private Sub cmdActionStep2_Click()
   ' perform action step 2
   MsgBox "Add your App function here.", vbInformation + vbOKOnly, "To be implemented"
End Sub

Private Sub cmdActionStep3_Click()
   ' perform action step 3
   MsgBox "Add your App function here.", vbInformation + vbOKOnly, "To be implemented"
End Sub

Private Sub cmdActionStep4_Click()
   ' perform action step 4
   MsgBox "Add your App function here.", vbInformation + vbOKOnly, "To be implemented"
End Sub
