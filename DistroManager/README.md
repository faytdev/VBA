# Distro Manager

Tons of bugs. Needs refactored. But it "works".

Used to manage distro lists. Add, Remove, and Edit Groups and Members. Create Named Ranges for each Group.

Use Launch Module to ensure required sheet is present. Run LaunchForm Sub. 

    Sub LaunchForm()
        If DataSheetCheck = False Then
            If CreateDataSheet = True Then
                DistroManager.Show
            End If
        Else
            DistroManager.Show
        End If
    End Sub

Else Create A Sheet Call "_DistroManager-DataSheet". 
Current set up uses headers:
_DistroManager-DataSheet Range A1 to = Header Value if you Chose
_DistroManager-DataSheet Range B1 to = Header Value if you Chose

Create At least One Group name in _DistroManager-DataSheet Range A2
Create At least One member in _DistroManager-DataSheet Range B2. Seperate each email with a semicolon
