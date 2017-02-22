Imports Autodesk.Revit
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Utility
'Copyright (c) 2017 Danny Polkinhorn

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

Imports Autodesk.Revit.UI.Events

<Transaction(TransactionMode.Manual)> _
<Regeneration(RegenerationOption.Manual)> _
Public Class UI
    Implements IExternalApplication

    Public Function OnShutdown(application As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
        'Clean up our dialog handler

        Return Result.Succeeded
    End Function

    Public Function OnStartup(application As UIControlledApplication) As Result Implements IExternalApplication.OnStartup

        'Set up the button in the ribbon
        Dim path As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim caption As String = "Crash Revit"
        Dim d As New PushButtonData(caption, caption, path, "CrashRevit.Commands")
        d.AvailabilityClassName = "CrashRevit.Availability"
        Dim p As RibbonPanel = application.CreateRibbonPanel(caption)
        Dim b As PushButton = TryCast(p.AddItem(d), PushButton)
        b.ToolTip = "Safely crashes Revit with an option to save the document."

        Return Result.Succeeded

    End Function

End Class

Public Class Availability
    Implements IExternalCommandAvailability
    Public Function IsCommandAvailable1(applicationData As UIApplication, selectedCategories As CategorySet) As Boolean Implements IExternalCommandAvailability.IsCommandAvailable
        Return True
    End Function
End Class