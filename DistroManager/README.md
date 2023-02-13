# Distro Manager

Used to manage distro lists. Add, Remove, and Edit Groups and Members. Create Named Ranges for each Group.

Use Launch Module to ensure required sheet is present.

'Sub LaunchForm()
    If DataSheetCheck = False Then
        If CreateDataSheet = True Then
            DistroManager.Show
        End If
    Else
        DistroManager.Show
    End If
End Sub'

Tons of bugs. Needs refactored. But it "works".