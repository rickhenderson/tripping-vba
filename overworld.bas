Option Explicit

' ----------------------------------------------------------------------
' OverWizard - A GUI for Overviewer to create Minecraft Maps
'
' Created By: Rick Henderson, @rickhenderson
' Created On: April 30, 2013
'
' Change log:
' Started: April 30: Version .001: Started code to run overviewer as
'   from Shell and added a custom tab using RibbonX.
'
' May 15, 2013: Version .002: Added better icons for the ribbon.
'   Added ability to browse to the Overviewer install folder and the output folder.
'
' ----------------------------------------------------------------------

'Callback for btnStart onAction
Sub startOW(control As IRibbonControl)
    ' Display the main form to start the application
    frmOWStart.Show
End Sub

'Callback for btnOverviewer onAction
Sub aboutOverviewer(control As IRibbonControl)
    MsgBox "Visit http://overviewer.org/", vbOKOnly & vbInformation, "OverWizard"
End Sub

'Callback for btnMC onAction
Sub aboutMC(control As IRibbonControl)
    ' Display info about Minecraft or open a site
    MsgBox "Visit http://www.minecraftwiki.net/", vbOKOnly & vbInformation, "Minecraft Wiki"
End Sub

