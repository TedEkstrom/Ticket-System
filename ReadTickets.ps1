Add-Type -AssemblyName PresentationFramework

## Varibles

$loadedtickets = @{}
$global:LoadedTicket = @()
$global:LastSelectTicket = ""
$UnixTime = [DateTimeOffset]::Now.ToUnixTimeSeconds()
$temp = [String]$UnixTime
$UnixTime = $UnixTime * [bigint]::Pow(10, (10-$temp.Length))
$Tag = $env:COMPUTERNAME

$Global:Settings = "$env:USERPROFILE\Readtickets" 
$Global:Path = ""

$Global:ticketOwner = ""
 
$newTickets = "$Global:Path\new\"
$solvedTickets = "$Global:Path\solved\"
$NotsolvedTickets = "$Global:Path\notsolved\"
$deletedTickets = "$Global:Path\Deleted\"
$prio1 = "$Global:Path\prio1\"
$prio2 = "$Global:Path\prio2\"
$prio3 = "$Global:Path\prio3\"
$paus = "$Global:Path\pause\"

####################################################

## Create profile
if ( !$(Test-Path -Path "$Global:Settings\Userprofile.json" -ErrorAction SilentlyContinue) ) {

        $Global:Path = "$Global:Settings\tickets"

        New-Item -Path "$Global:Settings\Userprofile.json" -ItemType File -Force -ErrorAction SilentlyContinue
        
        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'TicketPath' -Value $Global:Path
        
        $item | ConvertTo-Json | Out-File -FilePath "$Global:Settings\Userprofile.json" -Force
}

## Loading profile
$Global:Path = (Get-Content -Path "$Global:Settings\userprofile.json" | ConvertFrom-Json).TicketPath
 
#$Global:Path = "\\10.0.1.10\hidden$\tickets_test"
    
$newTickets = "$Global:Path\new\"
$solvedTickets = "$Global:Path\solved\"
$NotsolvedTickets = "$Global:Path\notsolved\"
$deletedTickets = "$Global:Path\Deleted\"
$prio1 = "$Global:Path\prio1\"
$prio2 = "$Global:Path\prio2\"
$prio3 = "$Global:Path\prio3\"
$pause = "$Global:Path\pause\"

if ( !($Global:Path -eq $null) -and !(Test-Path -Path "$Global:Path\new" -ErrorAction SilentlyContinue ) ) {
    
    New-Item -Path $newTickets -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $solvedTickets -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $NotsolvedTickets -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $deletedTickets -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $prio1 -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $prio2 -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $prio3 -ItemType Directory -ErrorAction SilentlyContinue
    New-Item -Path $pause -ItemType Directory -ErrorAction SilentlyContinue
}
 
####################################################

$inputXML = @"
<Window x:Class="TicketSystem.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ticket System" Height="600" Width="1000" UseLayoutRounding="True">
    <Grid>
        <DockPanel LastChildFill="True">
            <!-- Meny -->
            <Menu DockPanel.Dock="Top">
                <MenuItem Header="File">
                    <MenuItem Header="📂 Select ticket path"/>
                    <MenuItem Name="exitM" Header="❌ Exit"/>
                </MenuItem>
                <MenuItem Header="Settings">
                    <MenuItem Name="settings" Icon="⚙" Header="Settings"/>
                </MenuItem>
            </Menu>

            <!-- Header -->
            <TextBlock DockPanel.Dock="Top" Name="TitleBlock" Text="Ticket System"
                       FontSize="24" FontWeight="Bold"
                       Padding="10" Background="#333"
                       Foreground="White" TextAlignment="Center"/>

            <!-- Sidebar -->
            <StackPanel DockPanel.Dock="Left" Width="200" Background="#EEE">
                <Button Content="New Ticket" Name="newTicketB" Margin="5" Padding="0"
                        Background="#007ACC" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Rename Ticket" Name="renameTicketB" Margin="5" Padding="0"
                        Background="#007ACC" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Assign ticket" Name="assignTicketB" Margin="5" Padding="0"
                        Background="#FF1C4C6D" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Reset assign ticket" Name="resetAssignTicketB" Margin="5" Padding="0"
                        Background="#FF1C4C6D" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Choose prio 1" Name="choosePrio1B" Margin="5" Padding="0"
                        Background="#FF0B1E2B" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Choose prio 2" Name="choosePrio2B" Margin="5" Padding="0"
                        Background="#FF0B1E2B" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Choose prio 3" Name="choosePrio3B" Margin="5" Padding="0"
                        Background="#FF0B1E2B" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Pause ticket" Name="pauseTicketB" Margin="5" Padding="0"
                        Background="#FF62666F" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <Button Content="Delete ticket" Name="deleteTicketB" Margin="5" Padding="0"
                        Background="#FF754C4C" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
            </StackPanel>

            <!-- Användarval -->
            <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="10">
                <TextBlock Text="User: ...." Name="ticketOwnerL" VerticalAlignment="Center" Margin="5"/>
                <Button Content="Refresh Window" Name="refreshB" Margin="300,0,0,0" Padding="0"
                        Background="#FF0B1E2B" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="20"
                        HorizontalAlignment="Left"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
            </StackPanel>
            
            
            <!-- Ticket Filters -->
            <StackPanel DockPanel.Dock="Top" Background="{DynamicResource {x:Static SystemColors.ActiveCaptionBrushKey}}" Margin="10">
                <WrapPanel Background="{DynamicResource {x:Static SystemColors.ActiveCaptionBrushKey}}">
                    <CheckBox Name="newTicketsR" Content="New tickets" Margin="5"/>
                    <CheckBox Name="prio1R" Content="Prio 1" Margin="5"/>
                    <CheckBox Name="prio2R" Content="Prio 2" Margin="5"/>
                    <CheckBox Name="prio3R" Content="Prio 3" Margin="5"/>
                    <CheckBox Name="solvedR" Content="Solved" Margin="5"/>
                    <CheckBox Name="notSolvedR" Content="Not solved" Margin="5"/>
                    <CheckBox Name="pausedR" Content="Paused" Margin="5"/>
                    <CheckBox Name="withOutOwnerR" Content="Show assigned tickets (all tickets)" Margin="5"/>
                </WrapPanel>
            </StackPanel>

            <!-- Ticket List -->
            <ListView Name="ticketListViewT" Foreground="Black">
                <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                    </Style>
                </ListView.ItemContainerStyle>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Priority" Width="70">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Priority}" TextAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Assigned to" Width="70">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding AssignedTo}" TextAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Ticket Name" Width="300">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ticketName}" TextAlignment="Left"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="120">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Reported by" Width="70">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ReportedBy}" TextAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Date" Width="70">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Date}" TextAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                    </GridView>
                </ListView.View>
            </ListView>
        </DockPanel>
    </Grid>
</Window>
"@

#create window
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[xml]$XAML = $inputXML
#Read XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $MainWindow = [Windows.Markup.XamlReader]::Load( $reader )

    $settingsM = $MainWindow.FindName("settings")   
    $exitM = $MainWindow.FindName("exitM") 

    $ticketOwnerL = $MainWindow.FindName("ticketOwnerL")
    $tickets = $MainWindow.FindName("ticketListViewT")
    $newR = $MainWindow.FindName("newTicketsR")
    $prio1R = $MainWindow.FindName("prio1R")
    $prio2R = $MainWindow.FindName("prio2R")
    $prio3R = $MainWindow.FindName("prio3R")
    $notsolvedR = $MainWindow.FindName("notSolvedR") #< ändra till $notSolvedR
    $solvedR = $MainWindow.FindName("solvedR") #< ändra till $solvedR
    $pauseR = $MainWindow.FindName("pausedR")
    $withOutOwnerR = $MainWindow.FindName("withOutOwnerR")

    $newTicketB = $MainWindow.FindName("newTicketB")
    $renameTicketB = $MainWindow.FindName("renameTicketB")
    $assignticketB = $MainWindow.FindName("assignTicketB")
    $resetAssignTicketB = $MainWindow.FindName("resetAssignTicketB")
    $choosePrio1B = $MainWindow.FindName("choosePrio1B")
    $choosePrio2B = $MainWindow.FindName("choosePrio2B")
    $choosePrio3B = $MainWindow.FindName("choosePrio3B")
    $pauseTicketB = $MainWindow.FindName("pauseTicketB")
    $deleteTicketB = $MainWindow.FindName("deleteTicketB")
    $refreshB = $MainWindow.FindName("refreshB")
}
catch {
    Write-Warning $_.Exception
    throw
}

####################################################
## Load settings
function loadAutosaveSettings () {

    $AutoSave = Get-Content -Path "$Global:Settings\userprofile.json" | ConvertFrom-Json

    $NewR.IsChecked = $AutoSave.NewR
    $Prio1R.IsChecked = $AutoSave.Prio1R
    $Prio2R.IsChecked = $AutoSave.Prio2R
    $Prio3R.IsChecked = $AutoSave.Prio3R
    $SolvedR.IsChecked = $AutoSave.SolvedR
    $NotSolvedR.IsChecked = $AutoSave.NotSolvedR
    $WithOutOwnerR.IsChecked = $AutoSave.checked
    $Global:ticketOwner = $AutoSave.user
    
} loadAutosaveSettings

function autosaveSettings () {

     $item = New-Object PSObject
     $item | Add-Member -type NoteProperty -Name 'TicketPath' -Value $Global:Path
     $item | Add-Member -type NoteProperty -Name 'NewR' -Value $NewR.Checked
     $item | Add-Member -type NoteProperty -Name 'Prio1R' -Value $Prio1R.Checked
     $item | Add-Member -type NoteProperty -Name 'Prio2R' -Value $Prio2R.Checked
     $item | Add-Member -type NoteProperty -Name 'Prio3R' -Value $Prio3R.Checked
     $item | Add-Member -type NoteProperty -Name 'SolvdR' -Value $SolvdR.Checked
     $item | Add-Member -type NoteProperty -Name 'NotSolvdR' -Value $NotSolvdR.Checked
     $item | Add-Member -type NoteProperty -Name 'WithOutOwner' -Value $WithOutOwnerR.Checked
     $item | Add-Member -type NoteProperty -Name 'user' -Value $Global:ticketOwner
        
     $item | ConvertTo-Json | Out-File -FilePath "$Global:Settings\Userprofile.json"
}


## Laddar in användare 
$ticketOwnerL.Text = "User: $Global:ticketOwner"

####################################################

## Functions

function searchForTickets () {

    $tickets.Items.Clear()
    $loadedtickets.Clear()

    ## Status kan vara att de Väntar, Öppnad, osv ...


    if ( $NewR.IsChecked  ) {
        
       ## Eftersom sköningen måste genomföras på samtliga i-checkade,
       ## är det bättre att köra sökningen på samtliga.

       $NewT = (Get-ChildItem -Path $newTickets -File).FullName

       if ( $NewT ) {
        
          $NewT | ForEach-Object {
             $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
             $json = Get-Content -Path $_ | ConvertFrom-Json

             if ( $temp -like "*$($SearchTB.Text)*" ) {
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = "New"
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)

             } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = "New"
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)
             }
          }
       }
    }

    if ( $Prio1R.IsChecked ) {
       
       $Prio1T = (Get-ChildItem -Path $prio1 -File).FullName

       if ( $Prio1T ) {  
          $Prio1T | ForEach-Object {

            $isTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner

              if ( ($ticketOwnerL.Text -eq $null) -or $WithOutOwnerR.IsChecked  -or ($isTicketOwner -eq $null) ) {

                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json

                     if ( $temp -like "*$($SearchTB.Text)*" ) {
                         
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $tickets.HorizontalContentAlignment = "Center"
                         $loadedtickets.Add($temp, $_)

                     } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $loadedtickets.Add($temp, $_)
                     }
               } else {
                    if ( $Global:ticketOwner -eq $isTicketOwner ) {
                        
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
            
                         if ( $temp -like "*$($SearchTB.Text)*" ) {
                             
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)   
                             $loadedtickets.Add($temp, $_)

                         } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)   
                             $loadedtickets.Add($temp, $_)
                         }
                    }
               }
          }
       }
    }
    
    if ( $Prio2R.IsChecked  ) {
       
       $Prio2T = (Get-ChildItem -Path $prio2 -File).FullName

       if ( $Prio2T ) {  
          $Prio2T | ForEach-Object {
                
            $isTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner

              if ( ($ticketOwnerL.Text -eq $null) -or $WithOutOwnerR.IsChecked  -or ($isTicketOwner -eq $null) ) {

                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json

                     if ( $temp -like "*$($SearchTB.Text)*" ) {
                         
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $tickets.HorizontalContentAlignment = "Center"
                         $loadedtickets.Add($temp, $_)

                     } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $loadedtickets.Add($temp, $_)
                     }
               } else {
                    if ( $Global:ticketOwner -eq $isTicketOwner ) {
                        
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
            
                         if ( $temp -like "*$($SearchTB.Text)*" ) {
                             
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)   
                             $loadedtickets.Add($temp, $_)

                         } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)    
                             $loadedtickets.Add($temp, $_)
                         }
                    }
               }
          }
       }
    }

    if ( $Prio3R.IsChecked  ) {

       $Prio3T = (Get-ChildItem -Path $prio3 -File).FullName

       if ( $Prio3T ) {  
          $Prio3T | ForEach-Object {
             
            $isTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner

              if ( ($ticketOwnerL.Text -eq $null) -or $WithOutOwnerR.IsChecked  -or ($isTicketOwner -eq $null) ) {

                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json

                     if ( $temp -like "*$($SearchTB.Text)*" ) {
                         
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $tickets.HorizontalContentAlignment = "Center"
                         $loadedtickets.Add($temp, $_)

                     } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                         $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                         [void]$tickets.Items.Add($ticket)   
                         $loadedtickets.Add($temp, $_)
                     }
               } else {
                    if ( $Global:ticketOwner -eq $isTicketOwner ) {
                        
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
            
                         if ( $temp -like "*$($SearchTB.Text)*" ) {
                             
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)   
                             $loadedtickets.Add($temp, $_)

                         } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                             $ticket = New-Object PSObject -Property @{
                                ticketName = $temp
                                Status = $json.status
                                Priority = $json.prio
                                ReportedBy = $json.username
                                Date = $json.date
                                AssignedTO = $json.ticketOwner
                             }
                             [void]$tickets.Items.Add($ticket)   
                             $loadedtickets.Add($temp, $_)
                         }
                    }
               }
          }
       }
    }

    if ( $SolvdR.IsChecked  ) {
       $SolvdT = (Get-ChildItem -Path $solvdTickets -File).FullName

       if ( $SolvdT ) {  
          $SolvdT | ForEach-Object {
             $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
             $json = Get-Content -Path $_ | ConvertFrom-Json

             if ( $temp -like "*$($SearchTB.Text)*" ) {
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)

             } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)
             }
          }
       }
    }

    if ( $notSolvdR.IsChecked  ) {

       $NotSolvdT = (Get-ChildItem -Path $NotSolvdTickets -File).FullName

       if ( $NotSolvdT ) {  
          $NotSolvdT | ForEach-Object {
             $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
             $json = Get-Content -Path $_ | ConvertFrom-Json

             if ( $temp -like "*$($SearchTB.Text)*" ) {
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)

             } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)
             }
          }
       }
    }

    
    if ( $pauseR.IsChecked  ) {

       $pauseT = (Get-ChildItem -Path $pause -File).FullName

       if ( $pauseT ) {  
          $pauseT | ForEach-Object {
             $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
             $json = Get-Content -Path $_ | ConvertFrom-Json

             if ( $temp -like "*$($SearchTB.Text)*" ) {
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)

             } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
                    
                 $ticket = New-Object PSObject -Property @{
                            ticketName = $temp
                            Status = $json.status
                            Priority = $json.prio
                            ReportedBy = $json.username
                            Date = $json.date
                            AssignedTO = $json.ticketOwner
                         }
                 [void]$tickets.Items.Add($ticket)  
                 $loadedtickets.Add($temp, $_)
             }
          }
       }
    }
}
searchForTickets

function saveChanges () {
    # Change the json-file with new priority
    $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]
    $global:LoadedTicket | ConvertTo-Json | Out-File -FilePath $fileToWrite
}

## Ska implementeras i OpenTicket fönstret 
function solvdTicket () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {     
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $solvdTickets
        searchForTickets
    }
}

## Kanske ska implementeras i OpenTicket fönstret 
function NotSolvdTicket () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {        
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $NotSolvdTickets    
        searchForTickets
    }
}

function prio1 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {       
        $global:loadedticket.Prio = "Prio 1"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $prio1        
        searchForTickets
    }
}

function prio2 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {       
        $global:loadedticket.Prio = "Prio 2"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $prio2  
        searchForTickets
    }
}

function prio3 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {        
        $global:loadedticket.Prio = "Prio 3"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $prio3
        searchForTickets
    }
}

function pause () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $pause
        searchForTickets
    }
}

function deleteTicket () {

    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {      
        #openTicket
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$deletedTickets$($global:LastSelectTicket.ticketName.Replace(' ',''))_$(Get-Date -Format 'yyyyMMdd').json"
        searchForTickets
    }
}


function openTicket () {

$inputXML = @"
<Window x:Class="TicketSystem.OpenTicket"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:OpenTicket"
        mc:Ignorable="d"
        Title="Open ticket" Height="550" Width="800">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>
                <TextBlock Name="ticketNameT" Text="[ No name ]" FontSize="20" FontWeight="Bold" Foreground="#2C3E50" />
                <TextBlock Name="tagT" Text="Tag: Missing..." FontSize="10"  Padding="0,10,0,10" />
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <TextBlock Text="Status:" FontSize="16" Foreground="#27AE60" />
                    <ComboBox Name="statusCB"  Width="150" Margin="10,0,0,0">
                        <ComboBoxItem Content="In progress" IsSelected="True"/>
                        <ComboBoxItem Content="Awaiting response"/>
                        <ComboBoxItem Content="Resolved"/>
                        <ComboBoxItem Content="Closed"/>
                        <ComboBoxItem Content="Paused"/>
                        <ComboBoxItem Content="Escalated"/>
                    </ComboBox>
                </StackPanel>
                <TextBlock Name="priorityT" Text="Priority: Missing..." FontSize="16" Foreground="#F39C12" Padding="0,10,0,10" />
             
                <!-- Gör beskrivningen till en TextBox för mer utrymme och rullning -->
                <TextBox Name="descriptionT" Text="Missing a description.." 
                         FontSize="14" Foreground="Black"
                         TextWrapping="Wrap" AcceptsReturn="True" Height="68" VerticalScrollBarVisibility="Auto"
                         IsReadOnly="True" BorderBrush="Transparent" Background="Transparent"/>
                <Label Content="Update ticket continuously"/>

                <!-- TextBox för användarinmatning -->
                <TextBox Name="updateT" FontSize="15" Height="174" Margin="10,10,10,0" 
                     Text="Missing update..." Foreground="Black"
                     AcceptsReturn="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto"/>

                <!-- Flyttade knapparna längst ner -->
                <StackPanel Orientation="Horizontal" Margin="10">
                    <Button Name="updateB" Content="Update" Width="120" Background="#3498DB" Foreground="White" Padding="5" />
                    <Button Name="solvedB" Content="Solved" Width="120" Background="#2ECC71" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                    <Button Name="closeB"  Content="Close" Width="120" Background="#95A5A6" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                    <Button Name="notSolvableB"  Content="Not solvable" Width="120" Background="Darkred" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

    #create window
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $descriptionT = $Window.FindName("descriptionT")
        $ticketNameT = $Window.FindName("ticketNameT")
        $priorityT = $Window.FindName("priorityT")
        $statusCB = $Window.FindName("statusCB")
        $updateT = $Window.FindName("updateT")
        $tagT = $Window.FindName("tagT")
        $updateB = $Window.FindName("updateB")
        $solvedB = $Window.FindName("solvedB")
        $closeB = $Window.FindName("closeB")
        $notSolvableB = $Window.FindName("notSolvableB")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }
  
    $ticketNameT.Text = $loadedticket.title
    $priorityT.Text = "Priority: "+$loadedticket.prio

    $descriptionT.Text = $global:LoadedTicket.Error

    
    $desiredStatus = $Global:loadedticket.status
    foreach ($item in $statusCB.Items) {
        if ($item.Content -eq $desiredStatus) {
            $statusCB.SelectedItem = $item
            break
        }
    }
   
    $tagT.Text = "Tag: "+$global:LoadedTicket.tag

    if ( $loadedticket.prio -eq "Prio 1" ) {
        $priorityT.Foreground.Color = "Darkred"
    } elseif ( $loadedticket.prio -eq "Prio 2" ) {
        $priorityT.Foreground.Color = "#F39C12"
    } else {
        $priorityT.Foreground.Color = "Darkblue"
    }
    
    if ( $global:LoadedTicket.update.Length -gt 0 ) { 
        $updateT.Text = $global:LoadedTicket.update + "`rUppdatering, $(Get-Date -Format 'dd MMM yyyy')`rEnter Response here -->"
    } else {
        $updateT.Text = $global:LoadedTicket.update + "Uppdatering, $(Get-Date -Format 'dd MMM yyyy')`rEnter Response here -->"

    }

    $updateT.Dispatcher.InvokeAsync({
        $updateT.ScrollToEnd()
    }, [System.Windows.Threading.DispatcherPriority]::Background)

    $updateB.Add_Click({ 
        
        $temp = $updateT.Text
        $temp2 = $global:LoadedTicket.update + "`rUppdatering, $(Get-Date -Format 'dd MMM yyyy')`rEnter Response here -->"

        if ( $updateT.Text -eq $temp2 ) {
            
            $temp = $updateT.Text += "`r--------------------------------------------------------------------" 
               
        } else {

            $temp.replace("`rUppdatering, $(Get-Date -Format 'dd MMM yyyy')`r`n Enter Response here -->","")
        }

        $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]

        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'Title' -Value $global:LoadedTicket.Title
        $item | Add-Member -type NoteProperty -Name 'Computer' -Value $global:LoadedTicket.Computer
        $item | Add-Member -type NoteProperty -Name 'Tag' -Value $global:LoadedTicket.Tag
        $item | Add-Member -type NoteProperty -Name 'Date' -Value $global:LoadedTicket.Date
        $item | Add-Member -type NoteProperty -Name 'Error' -Value $global:LoadedTicket.Error
        $item | Add-Member -type NoteProperty -Name 'Name' -Value $global:LoadedTicket.Name
        $item | Add-Member -type NoteProperty -Name 'Update' -Value $temp
        $item | Add-Member -type NoteProperty -Name 'Username' -Value $global:LoadedTicket.Username
        $item | Add-Member -type NoteProperty -Name 'Prio' -Value $global:LoadedTicket.Prio
        $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusCB.SelectionBoxItem

        $global:loadedticket =  $item

        saveChanges
        $Window.Hide()
        searchForTickets
    })
    
    $solvedB.Add_Click({ 
        $global:LoadedTicket.status = $statusCB.SelectionBoxItem
        $global:loadedticket.Update += $updateT.Text + "`r`n-------------------------------- Solved ----------------------------`r`n"
        $global:loadedticket.Update += $updateT.Text + "`r`n--------------------------------------------------------------------`r`n"
        saveChanges
        $Window.Hide()
        solvdTicket
        searchForTickets
    })

    $closeB.Add_Click({$Window.Hide()})

    $notSolvableB.Add_Click({
    $global:LoadedTicket.status = $statusCB.SelectionBoxItem
        $global:loadedticket.Update += $updateT.Text + "`r`n----------------------- Not able to be Solved ----------------------`r`n"
        $global:loadedticket.Update += $updateT.Text + "`r`n--------------------------------------------------------------------`r`n"
        saveChanges
        $Window.Hide()
        NotSolvdTicket
        searchForTickets
    })

    [Void]$Window.ShowDialog();
}


function newTicket () {

$inputXML = @"
<Window x:Class="TicketSystem.NewTicket"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New Ticket" Height="550" Width="700"
        Background="#F0F0F0">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>
                <TextBlock Text="Create New Ticket" FontSize="20" FontWeight="Bold" Margin="0,0,0,20" Foreground="#2C3E50" />

                <TextBlock Text="Issue:" FontSize="16" Foreground="Black" />
                <TextBox Name="issueT" Height="30" Margin="0,5,0,10" Text="Enter issue description..." Foreground="Gray"/>

                <TextBlock Text="Priority:" FontSize="16" Foreground="Black" />
                <ComboBox Name="prioCB" Height="30" Margin="0,5,0,10">
                    <ComboBoxItem Content="Non" IsSelected="True"/>
                    <ComboBoxItem Content="Prio 1"/>
                    <ComboBoxItem Content="Prio 2"/>
                    <ComboBoxItem Content="Prio 3"/>
                </ComboBox>

                <TextBlock Text="Description:" FontSize="16" Foreground="#555555" />
                <TextBox Name="descriptionT" Height="120" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" Foreground="Gray"
                         Margin="0,5,0,10" Text="Enter detailed description..." />
                
                <TextBlock Text="Name:" FontSize="16" Foreground="Black" />
                <TextBox  Name="userT" Height="30" Margin="0,5,0,10" Text="Enter your name..." Foreground="Gray"/>

                <!-- Knapp för att skapa ticket -->
                <StackPanel Orientation="Horizontal" Margin="10" >
                    <Button Name="NewTicketB" Content="New ticket" Width="120" Background="#3498DB" Foreground="White" Padding="5"/>
                    <Button Name="closeB" Content="Close" Width="120" Background="#95A5A6" Foreground="White" Padding="5" Margin="5,0,0,0" />
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

    #create window
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $issueT = $Window.FindName("issueT")
        $prioCB = $Window.FindName("prioCB")
        $descriptionT = $Window.FindName("descriptionT")
        $NewTicketB = $Window.FindName("NewTicketB")
        $closeB = $Window.FindName("closeB")
        $userT = $Window.FindName("userT")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $issueT.Add_GotFocus({ $issueT.Text = "" })
    $descriptionT.Add_GotFocus({ $descriptionT.Text = "" })
    $issueT.Add_GotFocus({ $userT.Text = "" })

    $NewTicketB.Add_Click({
        if ( $issueT.Text -ne "Enter issue description..." -and $descriptionT.Text -ne "Enter detailed description..." -and $userT.Text -ne "Enter your name..." `
         -and $issueT.Text -ne "" -and $descriptionT.Text -ne "" -and $userT.Text -ne "" ) {

            $filtertitle = $issueT.Text.Replace(":", "-")

            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'Title' -Value $issueT.Text
            $item | Add-Member -type NoteProperty -Name 'Computer' -Value ""
            $item | Add-Member -type NoteProperty -Name 'Tag' -Value $env:COMPUTERNAME 
            $item | Add-Member -type NoteProperty -Name 'Date' -Value $(Get-Date -Format yyMMdd)
            $item | Add-Member -type NoteProperty -Name 'Error' -Value $descriptionT.Text
            $item | Add-Member -type NoteProperty -Name 'Name' -Value $userT.Text
            $item | Add-Member -type NoteProperty -Name 'Update' -Value ""
            $item | Add-Member -type NoteProperty -Name 'Username' -Value $env:USERNAME
            $item | Add-Member -type NoteProperty -Name 'Prio' -Value $prioCB.SelectionBoxItem
            $item | Add-Member -type NoteProperty -Name 'Status' -Value ""
            
            if ( $prioCB.SelectionBoxItem -eq "Non" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$newTickets\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 1" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Prio1\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 2" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Prio1\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 3" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Prio3\$($filtertitle).json"
            } 
            searchForTickets
            $Window.Hide()
        } else {
        
            Write-Error "You need to fill in every box."
        }
    })

    $closeB.Add_Click({$Window.Hide()})

    [Void]$Window.ShowDialog()    
}

function renameTicket () {

$inputXML = @"
<Window x:Class="TicketSystem.RenameTicket"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:RenameTicket"
        mc:Ignorable="d"
        Title="Rename ticket" Height="170" Width="360">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>

                <!-- Användarval -->
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="New name:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="newNameT" TextWrapping="Wrap" Text="TextBox" Width="200" Margin="10,0,0,0"/>
                </StackPanel>

                <!-- Kontrollknappar -->
                <StackPanel Orientation="Horizontal" Margin="10">
                    <Button Name="renameB" Content="Rename" Width="120" Background="#3498DB" Foreground="White" Padding="5"/>
                    <Button Name="closeB" Content="Close" Width="120" Background="#95A5A6" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

    #create window
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $newNameT = $Window.FindName("newNameT")
        $renameB = $Window.FindName("renameB")
        $closeB = $Window.FindName("closeB")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $newNameT.Text = $global:LastSelectTicket.ticketName

    $renameB.Add_Click({
        
        ### Inte klar än!

        if ( $newNameT.Text ) {

            $NewT = (Get-ChildItem -Path $newTickets -File).FullName
            $SolvdT = (Get-ChildItem -Path $SolvdTickets -File).FullName
            $Prio1T = (Get-ChildItem -Path $prio1 -File).FullName
            $Prio2T = (Get-ChildItem -Path $prio2 -File).FullName
            $Prio3T = (Get-ChildItem -Path $prio3 -File).FullName
            $NotSolvdT = (Get-ChildItem -Path $NotSolvdTickets -File).FullName
            $DeletedT = (Get-ChildItem -Path $deletedTickets -File).FullName

            $temp = $loadedtickets[$global:LastSelectTicket.ticketName].Split("\") | Select-Object -Last 1
            $tempDate = ($temp.Split("_") | Select-Object -Last 1).Replace(".json","")

            $date = Get-Date -Format "yyMMdd"
            
            if ( "$NewT" -notlike "*$($newNameT.Text)*" -or `
                   "$SolvdT" -notlike "*$($newNameT.Text)*" -or ` 
                      "$Prio1T" -notlike "*$($newNameT.Text)*" -or `
                          "$Prio2T" -notlike "*$($newNameT.Text)*" -or ` 
                             "$Prio3T" -notlike "*$($newNameT.Text)*" -or ` 
                                  "$NotSolvdT" -notlike "*$($newNameT.Text)*" -or ` 
                                      "$DeletedT" -notlike "*$($newNameT.Text)*" ) { 
                
                Rename-Item -LiteralPath $Global:loadedtickets[$global:LastSelectTicket.ticketName] -NewName "$($newNameT.Text).json"
                
                #Lägg till att finns det en sökning använd det resultet när den hämtar uppdaterad fil.
                searchForTickets
                $Window.Hide()
            
            } elseif ( "$($newNameT.Text)" -notlike "*$tempDate*" ) {    
                Write-Host "Tyvärr är namnet redan upptaget men jag föreslår: $($newNameT.Text)_$date."
                $newNameT.Text = "$($newNameT.Text)_$tempDate"

            } else {
                
                Write-Host "The name allready exist."
            }
        } else {
            
            Write-Host "Can´t be empty."
        }

    })

    $closeB.Add_Click({$Window.Hide()})

    [Void]$Window.ShowDialog();
    
}


function autosaveSettings () {

    $item = New-Object PSObject
    $item | Add-Member -type NoteProperty -Name 'TicketPath' -Value $Global:Path
    $item | Add-Member -type NoteProperty -Name 'NewR' -Value $NewR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'Prio1R' -Value $Prio1R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'Prio2R' -Value $Prio2R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'Prio3R' -Value $Prio3R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'SolvdR' -Value $SolvdR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'NotSolvdR' -Value $NotSolvdR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'WithOutOwner' -Value $WithOutOwnerR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'user' -Value $Global:ticketOwner
        
    $item | ConvertTo-Json | Out-File -FilePath "$Global:Settings\Userprofile.json"
}

function settings () {

$inputXML = @"
<Window x:Class="TicketSystem.Settings"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:SelectUser"
        mc:Ignorable="d"
        Title="Settings" Height="215" Width="383">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="10">
            <StackPanel>

                <!-- Användarval -->
                <TextBlock Text="User for ticket system" />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Select User:" FontSize="14" FontWeight="Bold" />
                    <ComboBox Name="selectUserCB" Width="200" Margin="10,0,0,0"/>
                </StackPanel>
                <TextBlock Text="Path to all the tickets to load" />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Path to ticket:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="pathT" TextWrapping="Wrap" Text="TextBox" Width="200" Margin="10,0,0,0"/>
                </StackPanel>

                <!-- Kontrollknappar -->
                <StackPanel Orientation="Horizontal" Margin="10">
                    <Button Name="saveB" Content="Save" Width="120" Background="#3498DB" Foreground="White" Padding="5"/>
                    <Button Name="closeB" Content="Close" Width="120" Background="#95A5A6" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

    #create window
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $selectUserCB = $Window.FindName("selectUserCB")
        $pathT = $Window.FindName("pathT")
        $saveB = $Window.FindName("saveB")
        $closeB = $Window.FindName("closeB")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $ticketOwners = Get-Content -Path "$Global:Path\owners.json" | ConvertFrom-Json
    $ticketOwners = $ticketOwners.ticketOwners.split(",").trim()

    $ticketOwners | ForEach-Object { $selectUserCB.Items.add($_) }


    $desiredUser = (Get-Content -Path "$Global:Settings\userprofile.json" | ConvertFrom-Json).ticketOwner
    foreach ($item in $selectUserCB.Items) {
        if ($item.Content -eq $desiredUser) {
            $selectUserCB.SelectedItem = $item
            break
        }
    }

    $pathT.Text = $Global:Path 

    if (!$Global:Path) {
        $pathT.Text = "A path need to be added"
    }

    $temp = $pathT.Text 
    $pathT.Add_GotFocus({ $pathT.Text = "" })
    $Window.Add_MouseDown({ $pathT.Text = $temp })

    $saveB.Add_Click({

        $Global:ticketOwner = $selectUserCB.SelectedValue
        $ticketOwnerL.Text = "User: $Global:ticketOwner"
        $Global:Path = $pathT.Text
        autosaveSettings
        searchForTickets
        $Window.Hide()
    })

    $closeB.Add_Click({$Window.Hide()})
    [Void]$Window.ShowDialog();
}

## Eventhandling

$settingsM.add_Click({ settings })
$exitM.add_Click({ $MainWindow.Close() })

$newR.Add_Click({ searchForTickets })
$solvedR.Add_Click({ searchForTickets })
$notsolvedR.Add_Click({ searchForTickets })
$prio1R.Add_Click({ searchForTickets })
$prio2R.Add_Click({ searchForTickets })
$prio3R.Add_Click({ searchForTickets })
$pauseR.Add_Click({ searchForTickets })
$withOutOwnerR.Add_Click({ searchForTickets })

$newTicketB.Add_Click({ newTicket })
$renameTicketB.Add_Click({ renameTicket })
$assignticketB.Add_Click({ })
$resetAssignTicketB.Add_Click({ })
$choosePrio1B.Add_Click({ prio1 })
$ChoosePrio2B.Add_Click({ prio2 })
$ChoosePrio3B.Add_Click({ prio3 })
$pauseTicketB.Add_Click({ pause })
$deleteTicketB.Add_Click({ deleteTicket })
$refreshB.Add_Click({ 
    $MainWindow.Dispatcher.Invoke({
        $MainWindow.UpdateLayout()
    })
    searchForTickets 
})

$tickets.Add_SelectionChanged({
    param($sender, $e)
    $selectedItem = $sender.SelectedItem
    $global:LastSelectTicket = $selectedItem
    if ( $tickets.Items.Count -eq $loadedtickets.Count ) {
        $global:LoadedTicket = Get-Content -Path $loadedtickets[$selectedItem.ticketName] -ErrorAction SilentlyContinue | ConvertFrom-Json
    }
})

$tickets.Add_MouseDoubleClick({ openTicket })
$MainWindow.Add_Closing({autosaveSettings})

[Void]$MainWindow.ShowDialog()