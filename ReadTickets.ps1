Add-Type -AssemblyName PresentationFramework

## Varibles

$loadedtickets = @{}
$global:LoadedTicket = @()
$global:LastSelectTicket = ""

$global:loadedAutotickets = @{}
$global:LoadedAutoTicket = @()
$global:LastAutoSelectTicket = ""

$UnixTime = [DateTimeOffset]::Now.ToUnixTimeSeconds()
$temp = [String]$UnixTime
$UnixTime = $UnixTime * [bigint]::Pow(10, (10-$temp.Length))
$Tag = $env:COMPUTERNAME

$Global:dateTime = ""

$Global:Settings = "$env:USERPROFILE\Readtickets" 
$Global:Path = ""

$Global:ticketOwner = ""
$Global:autoMove = $true
$Global:first = 8
$Global:second = 4
$Global:third = 0

$chooseTicketOwner = $false

$Global:cancelImport = $false

####################################################

## Create profile
if ( !$(Test-Path -Path "$Global:Settings\Userprofile.json" -ErrorAction SilentlyContinue) ) {

    $Global:Path = "$Global:Settings\tickets"

    New-Item -Path "$Global:Settings\Userprofile.json" -ItemType File -Force -ErrorAction SilentlyContinue | Out-Null
    $item = New-Object PSObject
    $item | Add-Member -type NoteProperty -Name 'TicketPath' -Value $Global:Path
    $item | Add-Member -type NoteProperty -Name 'NewR' -Value $true
    $item | Add-Member -type NoteProperty -Name 'Prio1R' -Value $true
    $item | Add-Member -type NoteProperty -Name 'Prio2R' -Value $true
    $item | Add-Member -type NoteProperty -Name 'Prio3R' -Value $true
    $item | Add-Member -type NoteProperty -Name 'SolvedR' -Value $false
    $item | Add-Member -type NoteProperty -Name 'NotSolvedR' -Value $false
    $item | Add-Member -type NoteProperty -Name 'WithOutOwner' -Value $true
    $item | Add-Member -type NoteProperty -Name 'automove' -Value $true
    $item | Add-Member -type NoteProperty -Name 'user' -Value $null
    $item | Add-Member -type NoteProperty -Name 'showWithNoOwners' -Value $false
        
    $item | ConvertTo-Json | Out-File -FilePath "$Global:Settings\Userprofile.json"
}
 
####################################################

$inputXML = @"
<Window x:Class="TicketSystem.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ReadTickets" Height="600" Width="1200" UseLayoutRounding="True">
    <Grid>
        <DockPanel LastChildFill="True">
            <!-- Meny -->
            <Menu DockPanel.Dock="Top">
                <MenuItem Header="File">
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
                <Button Content="Schedule Ticket" Name="autoTicketB" Margin="5" Padding="0"
                        Background="#007ACC" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="180" Height="30"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
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
                <TextBlock Text="User: ...." Name="ticketOwnerL" Width="200" VerticalAlignment="Center" Margin="5"/>
                <Button Content="[ C ]" Name="cleareB" Margin="0,3,0,0" Padding="0"
                        Background="#FF754C4C" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="40" Height="20"
                        HorizontalAlignment="Left"
                        VerticalAlignment="Top"
                        Cursor="Hand"/>
                <TextBox Text="" Name="searchTB" Width="200" VerticalAlignment="Center"  Margin="5" />
                <Button Content="Search" Name="searchB" Margin="0,3,0,0" Padding="0"
                        Background="#FF0B1E2B" Foreground="White"
                        FontWeight="Bold" BorderBrush="Transparent"
                        Width="100" Height="20"
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
                    <CheckBox Name="showWithNoOwners" Content="Show no assigned" Margin="5"/>
                    <CheckBox Name="showAllTicketsR" Content="Show all tickets" Margin="5"/>
                </WrapPanel>
            </StackPanel>

            <!-- Ticket List -->
            <ListView Name="ticketListViewT" Foreground="Black" FontSize="15">
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
                                    <TextBlock Text="{Binding Priority}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Assigned to" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding AssignedTo}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Ticket Name" Width="300">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ticketName}" TextAlignment="Center" 
                                    ToolTip="{Binding ticketName}" ToolTipService.InitialShowDelay="10" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextAlignment="Center" 
                                    ToolTip="{Binding Status}" ToolTipService.InitialShowDelay="10" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Deadline" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Deadline}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Reported by" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ReportedBy}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Reported" Width="80">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Date}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Visible" Width="80">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Visible}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Recurrent" Width="80">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Recurrent}" TextAlignment="Center" />
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
    $Global:prio1R = $MainWindow.FindName("prio1R")
    $Global:prio2R = $MainWindow.FindName("prio2R")
    $Global:prio3R = $MainWindow.FindName("prio3R")
    $notsolvedR = $MainWindow.FindName("notSolvedR") #< ändra till $notSolvedR
    $solvedR = $MainWindow.FindName("solvedR") #< ändra till $solvedR
    $Global:pauseR = $MainWindow.FindName("pausedR")
    $showAllTicketsR = $MainWindow.FindName("showAllTicketsR")
    $showWithNoOwners = $MainWindow.FindName("showWithNoOwners")

    $newTicketB = $MainWindow.FindName("newTicketB")
    $renameTicketB = $MainWindow.FindName("renameTicketB")
    $assignticketB = $MainWindow.FindName("assignTicketB")
    $resetAssignTicketB = $MainWindow.FindName("resetAssignTicketB")
    $choosePrio1B = $MainWindow.FindName("choosePrio1B")
    $choosePrio2B = $MainWindow.FindName("choosePrio2B")
    $choosePrio3B = $MainWindow.FindName("choosePrio3B")
    $Global:pauseTicketB = $MainWindow.FindName("pauseTicketB")
    $deleteTicketB = $MainWindow.FindName("deleteTicketB")
    $searchTB = $MainWindow.FindName("searchTB")
    $searchB = $MainWindow.FindName("searchB")
    $cleareB = $MainWindow.FindName("cleareB")
    $autoTicketB = $MainWindow.FindName("autoTicketB")
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
    $Global:prio1R.IsChecked = $AutoSave.Prio1R
    $Global:prio2R.IsChecked = $AutoSave.Prio2R
    $Global:prio3R.IsChecked = $AutoSave.Prio3R
    $SolvedR.IsChecked = $AutoSave.SolvedR
    $NotSolvedR.IsChecked = $AutoSave.NotSolvedR
    $showAllTicketsR.IsChecked = $AutoSave.WithOutOwner
    $Global:autoMove = $AutoSave.automove
    $Global:ticketOwner = $AutoSave.user
    $showWithNoOwners.IsChecked = $AutoSave.showWithNoOwners
    $Global:first = $AutoSave.first
    $Global:second = $AutoSave.second
    $Global:third = $AutoSave.third
    $Global:Path = $AutoSave.ticketPath
    $Global:chooseTicketOwner = $AutoSave.chooseTicketOwner

    ## Maybe I'll improve this later. A bit of a hack.
    if ( $Global:chooseTicketOwner -eq $null ) {
        $Global:chooseTicketOwner = $false
    }

    if ( [string]::IsNullOrEmpty($Global:first) -or [string]::IsNullOrEmpty($Global:second) -or [string]::IsNullOrEmpty($Global:third) ) {
        $Global:first = 8
        $Global:second = 4
        $Global:third = 0
    }

    $Global:newTickets = "$Global:Path\new\"
    $Global:solvedTickets = "$Global:Path\solved\"
    $Global:NotsolvedTickets = "$Global:Path\notsolved\"
    $Global:deletedTickets = "$Global:Path\Deleted\"
    $Global:prio1 = "$Global:Path\prio1\"
    $Global:prio2 = "$Global:Path\prio2\"
    $Global:prio3 = "$Global:Path\prio3\"
    $Global:pause = "$Global:Path\pause\"
    $Global:autoTickets = "$Global:Path\autoTickets\"

    
    if ( !($Global:Path -eq $null) -and !(Test-Path -Path "$Global:Path\new" -ErrorAction SilentlyContinue ) ) {
    
        [void](New-Item -Path $Global:newTickets -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:solvedTickets -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:NotsolvedTickets -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:deletedTickets -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:prio1 -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:prio2 -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:prio3 -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:pause -ItemType Directory -ErrorAction SilentlyContinue)
        [void](New-Item -Path $Global:autoTickets -ItemType Directory -ErrorAction SilentlyContinue)
        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'ticketOwners' -Value 'User1'       
        $item | ConvertTo-Json | Out-File -FilePath "$Global:Path\owners.json"
    }
} 
loadAutosaveSettings

## Laddar in användare 
$ticketOwnerL.Text = "User: $Global:ticketOwner"

####################################################

## Functions

$global:isProgrssbar = $false

Function showProgressBar ( $show ) {
  
  if ( $show -and !$global:isProgrssbar ) {
    # Credit to https://tiberriver256.github.io/powershell/PowerShellProgress-Pt1/
    # Made some minor changes on the UI

        $global:isProgrssbar = $true

        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        $syncHash = [hashtable]::Synchronized(@{})
        $newRunspace =[runspacefactory]::CreateRunspace()
        $syncHash.Runspace = $newRunspace
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
        $PowerShellCommand = [PowerShell]::Create().AddScript({
        [xml]$xaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="Laddar..." Height="150" Width="300" WindowStartupLocation="CenterScreen" Topmost="True">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
            <TextBlock Text="Arbetar..." Margin="0,0,0,10" HorizontalAlignment="Center"/>
            <ProgressBar IsIndeterminate="True" Height="20" Width="200" HorizontalAlignment="Center"/>
        </StackPanel>
    </Window>
"@

        $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
        #===========================================================================
        # Store Form Objects In PowerShell
        #===========================================================================
        $xaml.SelectNodes("//*[@Name]") | %{ $SyncHash."$($_.Name)" = $SyncHash.Window.FindName($_.Name)}

        $syncHash.top
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error = $Error

    })
    $PowerShellCommand.Runspace = $newRunspace
    $data = $PowerShellCommand.BeginInvoke()


    Register-ObjectEvent -InputObject $SyncHash.Runspace `
            -EventName 'AvailabilityChanged' `
            -Action {

                    if($Sender.RunspaceAvailability -eq "Available")
                    {
                        $Sender.Closeasync()
                        $Sender.Dispose()
                    }
                }
    return $SyncHash
  }      
}

function closeProgressBar {
    
    Param (
        [Parameter(Mandatory=$false)]
        [System.Object[]]$ProgressBar,

        [Parameter()]
        $Show   
    )

    if ( $show -and $global:isProgrssbar ) {
        
        Start-Sleep -Milliseconds 50
        $global:isProgrssbar = $false
        $ProgressBar.Window.Dispatcher.Invoke([action]{
          $ProgressBar.Window.Close()
        }, "Normal")
    }

}


function scanJsonFiles ( $switch, $temp, $json) {
    
    if ( $switch -eq 0 ) {
        $ticket = New-Object PSObject -Property @{
            ticketName = $temp
            Status = "New"
            Priority = $json.prio
            ReportedBy = $json.username
            Date = $json.date
            AssignedTO = $json.ticketOwner
            DeadLine = $json.deadLine
        }
    } else {

        if ( [string]::IsNullOrEmpty($json.id) ) {

            $ticket = New-Object PSObject -Property @{
            ticketName = $temp
            Status = $json.status
            Priority = $json.prio
            ReportedBy = $json.username
            Date = $json.date
            AssignedTO = $json.ticketOwner
            DeadLine = $json.deadLine
            Visible = $json.visible
            }
        } else {
            $ticket = New-Object PSObject -Property @{
            ticketName = $temp
            Status = $json.status
            Priority = $json.prio
            ReportedBy = $json.username
            Date = $json.date
            AssignedTO = $json.ticketOwner
            DeadLine = $json.deadLine
            Visible = $json.visible
            Recurrent = "X"
            }
        }
    }

    if ( $temp -like "*$($SearchTB.Text)*" ) {
        [void]$tickets.Items.Add($ticket)  
        $loadedtickets.Add($temp, $_)

    } elseif ( $(Get-Content -Path $_) -like "*$($SearchTB.Text)*" ) {
            
        [void]$tickets.Items.Add($ticket)  
        $loadedtickets.Add($temp, $_)
    }
}

$errorSet = $false

function searchForTickets ($show) { 
    if ( ![string]::IsNullOrEmpty($Global:Path) ) {

        if ( (Test-Path -Path $Global:Path -ErrorAction SilentlyContinue) -and !$errorSet ) {
  
            $tickets.Items.Clear()
            $loadedtickets.Clear()

            $ProgressBar = showProgressBar -show $show
 
            if ( $NewR.IsChecked  ) {
     
               $NewT = (Get-ChildItem -Path $Global:newTickets -File).FullName

               if ( $NewT ) {
                  $NewT | ForEach-Object {
                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json
                     scanJsonFiles -switch 0 -temp $temp -json $json
                  }
               }
            }

            if ( $Global:prio1R.IsChecked ) {
       
               $Global:prio1T = (Get-ChildItem -Path $Global:prio1 -File).FullName

               if ( $Global:prio1T ) {  
                  $Global:prio1T | ForEach-Object {
                        $containsTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                        $json = Get-Content -Path $_ | ConvertFrom-Json
                
                        if ( $showAllTicketsR.IsChecked ) {

                            scanJsonFiles -switch 1 -temp $temp -json $json

                        } elseif ( $showWithNoOwners.IsChecked ) {
                    
                            if ( [string]::IsNullOrEmpty($containsTicketOwner) ) { 
                        
                                scanJsonFiles -switch 1 -temp $temp -json $json
                            }
                        } elseif ( !$showAllTicketsR.IsChecked ) {
                    
                            if ( ![string]::IsNullOrEmpty($containsTicketOwner) ) { 
                        
                                if ( $containsTicketOwner -eq $Global:ticketOwner ) {
                                    scanJsonFiles -switch 1 -temp $temp -json $json
                                }
                            }
                        }           
                    }
                }
            }
    
            if ( $Global:prio2R.IsChecked  ) {
                $Global:prio2T = (Get-ChildItem -Path $Global:prio2 -File).FullName
                if ( $Global:prio2T ) {  
                    $Global:prio2T | ForEach-Object {
                        $containsTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                        $json = Get-Content -Path $_ | ConvertFrom-Json
                
                        if ( $showAllTicketsR.IsChecked ) {

                            scanJsonFiles -switch 1 -temp $temp -json $json

                        } elseif ( $showWithNoOwners.IsChecked ) {
                    
                            if ( [string]::IsNullOrEmpty($containsTicketOwner) ) { 
                        
                                scanJsonFiles -switch 1 -temp $temp -json $json
                            }
                        } elseif ( !$showAllTicketsR.IsChecked ) {  
                    
                            if ( ![string]::IsNullOrEmpty($containsTicketOwner) ) { 
                                if ( $containsTicketOwner -eq $Global:ticketOwner ) {
                                    scanJsonFiles -switch 1 -temp $temp -json $json
                                }
                            }
                        }           
                    }
                }
            }

            if ( $Global:prio3R.IsChecked  ) {
                $Global:prio3T = (Get-ChildItem -Path $Global:prio3 -File).FullName       
                if ( $Global:prio3T ) {  
                    $Global:prio3T | ForEach-Object {
                        $containsTicketOwner = (Get-Content -Path $_ | ConvertFrom-Json).ticketOwner
                        $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                        $json = Get-Content -Path $_ | ConvertFrom-Json
                
                        if ( $showAllTicketsR.IsChecked ) {

                            scanJsonFiles -switch 1 -temp $temp -json $json

                        } elseif ( $showWithNoOwners.IsChecked ) {
                    
                            if ( [string]::IsNullOrEmpty($containsTicketOwner) ) { 
                        
                                scanJsonFiles -switch 1 -temp $temp -json $json
                            }
                        } elseif ( !$showAllTicketsR.IsChecked ) {  
                    
                            if ( ![string]::IsNullOrEmpty($containsTicketOwner) ) { 
                                if ( $containsTicketOwner -eq $Global:ticketOwner ) {
                                    scanJsonFiles -switch 1 -temp $temp -json $json
                                }
                            }
                        }           
                    }
                }
            }

            if ( $SolvedR.IsChecked  ) {
               $SolvedT = (Get-ChildItem -Path $Global:solvedTickets -File).FullName
               if ( $SolvedT ) {  
                  $SolvedT | ForEach-Object {
                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json
                     scanJsonFiles -switch 1 -temp $temp -json $json            
                  }
               }
            }

            if ( $notSolvedR.IsChecked  ) {
               $NotSolvedT = (Get-ChildItem -Path $Global:NotsolvedTickets -File).FullName
               if ( $NotSolvedT ) {  
                  $NotSolvedT | ForEach-Object {
                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json
                     scanJsonFiles -switch 1 -temp $temp -json $json 
                  }
               }
            }
    
            if ( $Global:pauseR.IsChecked  ) {
               $Global:pauseT = (Get-ChildItem -Path $Global:pause -File).FullName
               if ( $Global:pauseT ) {  
                  $Global:pauseT | ForEach-Object {
                     $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                     $json = Get-Content -Path $_ | ConvertFrom-Json
                     scanJsonFiles -switch 1 -temp $temp -json $json
                  }
               }
            }

            closeProgressBar -ProgressBar $ProgressBar -Show $show 

        } else {
        
            Write-Warning $Error
            $errorSet = $true
        }
    } else {
        
        Write-Warning "The path is empty, which it must not be."
    }
}
searchForTickets

function saveChanges () {
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) { 

        $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]
        $global:LoadedTicket | ConvertTo-Json | Out-File -FilePath $fileToWrite
    }
}


function solvedTicket () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) { 
        
        #Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$Global:solvedTickets$($global:LastSelectTicket.ticketName.Replace(' ',''))_$(Get-Date -Format 'yyyyMMdd_ss').json"
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$Global:solvedTickets$($global:LastSelectTicket.ticketName) (Solved-$(Get-Date -Format 'ddMMyy-hhmmss')).json"
        searchForTickets
    }
}

function NotsolvedTicket () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {  
    
        
        #Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$Global:NotsolvedTickets$($global:LastSelectTicket.ticketName.Replace(' ',''))_$(Get-Date -Format 'yyyyMMdd_ss').json"
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$Global:NotsolvedTickets$($global:LastSelectTicket.ticketName) (Not-solved-$(Get-Date -Format 'ddMMyy-hhmmss')).json"
        searchForTickets
    }
}

function prio1 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {       
        $global:loadedticket.Prio = "Prio 1"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $Global:prio1        
        searchForTickets
    }
}

function prio2 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0 ) {       
        $global:loadedticket.Prio = "Prio 2"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $Global:prio2  
        searchForTickets
    }
}

function prio3 () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {        
        $global:loadedticket.Prio = "Prio 3"
        saveChanges
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $Global:prio3
        searchForTickets
    }
}

function pause () {
    
    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination $Global:pause
        searchForTickets
    }
}

function deleteTicket () {

    if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {      
        Move-Item -Path $loadedtickets[$global:LastSelectTicket.ticketName] -Destination "$Global:deletedTickets$($global:LastSelectTicket.ticketName.Replace(' ',''))_$(Get-Date -Format 'yyyyMMdd_ss').json"
        searchForTickets
    }
}

function assignTicketOwner () {


$inputXML = @"
<Window x:Class="WpfApp1.assignTicket"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="Assign To Ticket" Height="180" Width="300">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="10">
            <StackPanel>

                <!-- Användarval -->
                <TextBlock Text="User for ticket system" />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Select User:" FontSize="14" FontWeight="Bold" />
                    <ComboBox Name="selectUserCB" Width="200" Margin="10,0,0,0"/>
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

    function assignOwner () {
           if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {

            $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]

            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'Title' -Value $global:LoadedTicket.Title
            $item | Add-Member -type NoteProperty -Name 'Computer' -Value $global:LoadedTicket.Computer
            $item | Add-Member -type NoteProperty -Name 'Tag' -Value $global:LoadedTicket.Tag
            $item | Add-Member -type NoteProperty -Name 'Date' -Value $global:LoadedTicket.Date
            $item | Add-Member -type NoteProperty -Name 'Error' -Value $global:LoadedTicket.Error
            $item | Add-Member -type NoteProperty -Name 'Name' -Value $global:LoadedTicket.Name
            $item | Add-Member -type NoteProperty -Name 'Update' -Value $global:LoadedTicket.Update
            $item | Add-Member -type NoteProperty -Name 'Username' -Value $global:LoadedTicket.Username
            $item | Add-Member -type NoteProperty -Name 'Prio' -Value $global:LoadedTicket.Prio

            if ( $Global:chooseTicketOwner ) {
                $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value $selectUserCB.Text
            } else {
                $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value $Global:ticketOwner
            }
            $item | Add-Member -type NoteProperty -Name 'Status' -Value $global:LoadedTicket.SelectedItem
            $item | Add-Member -type NoteProperty -Name 'id' -Value $global:LoadedTicket.id
            $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $global:LoadedTicket.deadLine
            $item | Add-Member -type NoteProperty -Name 'visible' -Value $global:LoadedTicket.visible

            $global:loadedticket =  $item

            $SelectedItem = $Tickets.SelectedItems.Text

            saveChanges
            searchForTickets
        } 
    }

    $saveB.Add_Click({
        assignOwner
        $Window.Hide()
    })

    $closeB.Add_Click({$Window.Hide()})
    
    if ( $Global:chooseTicketOwner ) {
        [Void]$Window.ShowDialog()
    } else {
        assignOwner
    }
}


function resetTicketOwner () {

     if ( $Tickets.SelectedItems.ticketName.Length -gt 0  ) {

        $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]

        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'Title' -Value $global:LoadedTicket.Title
        $item | Add-Member -type NoteProperty -Name 'Computer' -Value $global:LoadedTicket.Computer
        $item | Add-Member -type NoteProperty -Name 'Tag' -Value $global:LoadedTicket.Tag
        $item | Add-Member -type NoteProperty -Name 'Date' -Value $global:LoadedTicket.Date
        $item | Add-Member -type NoteProperty -Name 'Error' -Value $global:LoadedTicket.Error
        $item | Add-Member -type NoteProperty -Name 'Name' -Value $global:LoadedTicket.Name
        $item | Add-Member -type NoteProperty -Name 'Update' -Value $global:LoadedTicket.Update
        $item | Add-Member -type NoteProperty -Name 'Username' -Value $global:LoadedTicket.Username
        $item | Add-Member -type NoteProperty -Name 'Prio' -Value $global:LoadedTicket.Prio
        $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value ""
        $item | Add-Member -type NoteProperty -Name 'Status' -Value $global:LoadedTicket.Status

        $item | Add-Member -type NoteProperty -Name 'id' -Value $global:LoadedTicket.id

        $item | Add-Member -type NoteProperty -Name 'visible' -Value $global:LoadedTicket.visible

        $global:loadedticket =  $item

        $SelectedItem = $Tickets.SelectedItems.Text

        saveChanges
        searchForTickets
    } 
}

$Global:comment = ""

function addComment () {

$inputXML = @"
<Window x:Class="TicketSystem.Comment"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:OpenTicket"
        mc:Ignorable="d"
        Title="Add a comment" Height="550" Width="800">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>
                <TextBlock Text="Add a comment" FontSize="20" FontWeight="Bold" Foreground="#2C3E50" />

                <TextBox Name="commentT" FontSize="15" Height="360" Margin="10,20,10,0" 
                     Text="" Foreground="Black" 
                     AcceptsReturn="True"  TextWrapping="Wrap"
                     IsReadOnly="False" VerticalScrollBarVisibility="Auto"/>

                <!-- Flyttade knapparna längst ner -->
                <StackPanel Orientation="Horizontal" Margin="10">
                    <Button Name="addCommentB" Content="Add comment" Width="220" Background="#FF4DA3DC" Foreground="White" Padding="5" />
                    <Button Name="closeB"  Content="Close" Width="120" Background="Darkred" Foreground="White" Padding="5" Margin="5,0,0,0"/>
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
        $commentT = $Window.FindName("commentT")
        $addCommentB = $Window.FindName("addCommentB")
        $closeB = $Window.FindName("closeB")

    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $commentT.Text = "Update, $(Get-Date -Format 'dd MMM yyyy')"

    $addCommentB.Add_Click({

        ## Måste komma på hur jag ska koppla två delar med varandra.
        ## Det enklaste är nog bara att använda en ny global variabel för comments.
        
        if ( $allUpdatesT.Text -eq "" ) {
            $Global:comment = $allUpdatesT.Text + $commentT.Text + "`r--------------------------------------------------------------------" 
        } else {
            $Global:comment = $allUpdatesT.Text + "`r" + $commentT.Text + "`r--------------------------------------------------------------------"
        }
    
        $Window.hide()
    })


    $CloseB.Add_Click({ $Window.hide() })

    [Void]$Window.ShowDialog();
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
        Title="Open ticket" Height="820" Width="800">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>
                <TextBlock Name="ticketNameT" Text="[ No name ]" FontSize="20" FontWeight="Bold" Foreground="#2C3E50" />
                <TextBlock Name="tagT" Text="Tag: Missing..." FontSize="10" Padding="0,10,0,10" />

                <TextBlock Name="priorityT" Text="Priority: Missing..." FontSize="16" Foreground="#F39C12" Padding="0,0,0,0" />

                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <TextBlock Name="visibleT" Text="Visible: Private" FontSize="16"/>
                    <Button Name="visibleB" Content="Public" Width="60" FontSize="12" 
                         HorizontalAlignment="Left" Margin="10,0,0,0" Height="20" />
                </StackPanel>
                
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <TextBlock Text="Status:" FontSize="16" Foreground="#27AE60" />
                    <ComboBox Name="statusCB"  Width="150" Margin="10,0,15,0">
                        <ComboBoxItem Content="Not set"/>
                        <ComboBoxItem Content="In progress" IsSelected="True"/>
                        <ComboBoxItem Content="Awaiting response"/>
                        <ComboBoxItem Content="Pending"/>
                        <ComboBoxItem Content="Response"/>
                        <ComboBoxItem Content="External"/>
                        <ComboBoxItem Content="Resolved"/>
                        <ComboBoxItem Content="Reopened"/>
                        <ComboBoxItem Content="Rejected"/>
                        <ComboBoxItem Content="Closed without action"/>
                        <ComboBoxItem Content="Awaiting approval"/>
                        <ComboBoxItem Content="Planned"/>
                    </ComboBox>
                    <TextBox Name="statusT"   Text="Custom text" FontSize="15" Width="150" Margin="0,0,0,0" 
                     Foreground="Black" AcceptsReturn="True" TextWrapping="Wrap" Visibility="Hidden"
                     VerticalScrollBarVisibility="Auto"/>
                </StackPanel>

                <StackPanel Orientation="Horizontal" Margin="0,15,0,10">
                    <TextBlock Name="DeadlineT" Text="Deadline: Not set" FontSize="16" Padding="0,0,0,0" />
                    <Button Name="deadlineB" Content="Add deadline" Width="100" FontSize="12" 
                            HorizontalAlignment="Left" Margin="10,0,0,0" Height="20" />
                    <Button Name="resetDeadlineB" Content="Reset" Width="60" FontSize="12" 
                            HorizontalAlignment="Left" Margin="5,0,0,0" Height="20" />
                </StackPanel>
                <!-- Gör beskrivningen till en TextBox för mer utrymme och rullning -->
                <Label Content="Description"/>
                <TextBox Name="descriptionT" Text="Missing a description.."
                         FontSize="14" Foreground="Black" Width="705"
                         TextWrapping="Wrap" AcceptsReturn="True" Height="100" VerticalScrollBarVisibility="Auto"
                         IsReadOnly="True" Background="Transparent"/>

                <Label Name="allPriviusUpdatesL" Content="All privius updates"/>

                <!-- TextBox för användarinmatning -->

                <TextBox Name="allUpdatesT" FontSize="15" Height="300" Margin="10,0,10,0" 
                     Text="Missing update..." Foreground="Black" 
                     AcceptsReturn="True"  TextWrapping="Wrap"
                     IsReadOnly="True" VerticalScrollBarVisibility="Auto"/>

                <Button Name="editB" Content="Edit" Width="60" FontSize="12" 
                        Height="17" Background="#FFC59200" Foreground="White" Padding="0" 
                        HorizontalAlignment="Left" Margin="15,5,0,10"  />

                <!-- Flyttade knapparna längst ner -->
                <StackPanel Orientation="Horizontal" Margin="10">
                    <Button Name="addCommentB" Content="[ Add comment ]" Width="120" Background="#FF4DA3DC" Foreground="White" Padding="5" />
                    <Button Name="closeB"  Content="Close and update" Width="120" Background="#3498DB" Foreground="White" Padding="5" Margin="5,0,0,0"/>
                    <Button Name="solvedB" Content="Solved" Width="120" Background="#2ECC71" Foreground="White" Padding="5" Margin="5,0,0,0"/>
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
        $addCommentB = $Window.FindName("addCommentB")
        $descriptionT = $Window.FindName("descriptionT")
        $ticketNameT = $Window.FindName("ticketNameT")
        $priorityT = $Window.FindName("priorityT")
        $statusCB = $Window.FindName("statusCB")
        #$updateT = $Window.FindName("updateT")
        $allUpdatesT = $Window.FindName("allUpdatesT")
        $tagT = $Window.FindName("tagT")
        $editB = $Window.FindName("editB")
        $solvedB = $Window.FindName("solvedB")
        $closeB = $Window.FindName("closeB")
        $notSolvableB = $Window.FindName("notSolvableB")
        $tagBorder = $Window.FindName("tagBorder")
        $allPriviusUpdatesL = $Window.FindName("allPriviusUpdatesL")
        $statusT = $Window.FindName("statusT")
        $DeadlineT = $Window.FindName("DeadlineT")
        $DeadlineB = $Window.FindName("deadlineB")
        $resetDeadlineB = $Window.FindName("resetDeadlineB")

        $visibleB = $Window.FindName("visibleB")
        $visibleT = $Window.FindName("visibleT")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $Global:dateTime = ""
    $ticketNameT.Text = $loadedticket.title
    $priorityT.Text = "Priority: "+$loadedticket.prio
    $descriptionT.Text = $global:LoadedTicket.Error
    $desiredStatus = $Global:loadedticket.status
  

    $notSet = $true
    foreach ($item in $statusCB.Items) {
        if ($item.Content -eq $desiredStatus) {
            $statusCB.SelectedItem = $item
            $notSet = $false
            break
        } else {

            $statusCB.SelectedItem = "Not set"
        }
    }

    if ( $notSet ) {
        foreach ($item in $statusCB.Items) {
        if ($item.Content -eq "Not set") {
            $statusCB.SelectedItem = $item
            break
            } 
        }
        $statusT.Visibility = "Visible"
        $statusT.Text = $desiredStatus
    }
   
    $tagT.Text = "Tag: "+$global:LoadedTicket.tag

    if ( $loadedticket.prio -eq "Prio 1" ) {
        $priorityT.Foreground.Color = "Darkred"
    } elseif ( $loadedticket.prio -eq "Prio 2" ) {
        $priorityT.Foreground.Color = "#F39C12"
    } else {
        $priorityT.Foreground.Color = "Darkblue"
    }

    if ( $global:LoadedTicket.visible -eq "Public" ) {
        
        $visibleT.Text = "Visible: Public"
        $visibleB.Content = "Private"
    } else {

        $visibleT.Text = "Visible: Private"
        $visibleB.Content = "Public"
    }

    $allUpdatesT.Text = $global:LoadedTicket.update
    
    $allUpdatesT.Dispatcher.InvokeAsync({
        $allUpdatesT.ScrollToEnd()
    }, [System.Windows.Threading.DispatcherPriority]::Background)

    if ( !($global:LoadedTicket.deadLine -eq $null -or $global:LoadedTicket.deadLine -eq "") ) {
        $DeadlineT.Text = "Deadline: $($global:LoadedTicket.deadLine)"
    } else {
        $DeadlineT.Text = "Deadline: Not set"
    }

    $closeB.Add_Click({

        $fileToWrite = $global:loadedtickets[$global:LastSelectTicket.ticketName]

        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'Title' -Value $global:LoadedTicket.Title
        $item | Add-Member -type NoteProperty -Name 'Computer' -Value $global:LoadedTicket.Computer
        $item | Add-Member -type NoteProperty -Name 'Tag' -Value $global:LoadedTicket.Tag
        $item | Add-Member -type NoteProperty -Name 'Date' -Value $global:LoadedTicket.Date
        $item | Add-Member -type NoteProperty -Name 'Error' -Value $global:LoadedTicket.Error
        $item | Add-Member -type NoteProperty -Name 'Name' -Value $global:LoadedTicket.Name
        $item | Add-Member -type NoteProperty -Name 'Update' -Value $allUpdatesT.Text
        $item | Add-Member -type NoteProperty -Name 'Username' -Value $global:LoadedTicket.Username
        $item | Add-Member -type NoteProperty -Name 'Prio' -Value $global:LoadedTicket.Prio
        $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value $global:LoadedTicket.ticketOwner

        if ( !($statusCB.SelectionBoxItem -eq "Not set") ) {
            $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusCB.SelectionBoxItem     
        } else {
            $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusT.Text
        }

        if ( $DeadlineT.Text -like "*Not set" ) {
            $item | Add-Member -type NoteProperty -Name 'deadLine' -Value ""
        } else {
            $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $DeadlineT.Text.replace("Deadline: ", "")
        }

        $item | Add-Member -type NoteProperty -Name 'id' -Value $global:LoadedTicket.id
        
        $item | Add-Member -type NoteProperty -Name 'visible' -Value $visibleT.Text.Replace("Visible: ", "")

        $global:loadedticket =  $item


        saveChanges
        $Window.Hide()
        searchForTickets
    })

    
    $addCommentB.Add_Click({ 
        
        addComment
        $allUpdatesT.Text = $Global:comment
    })

    $editB.Add_Click({

        if ( $editB.Content -eq "Edit" ) {
            
            $allUpdatesT.IsReadOnly = $false
            $allUpdatesT.Foreground = "Black"
            $allUpdatesT.FontWeight = [System.Windows.FontWeights]::Bold
            $editB.Content = "Save"
            $editB.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})

        } else {
            $editB.Content = "Edit"
            $allUpdatesT.IsReadOnly = $true
            $editB.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})
            $allUpdatesT.Foreground = "Black"
            $allUpdatesT.FontWeight = [System.Windows.FontWeights]::Normal

            $allPriviusUpdatesL.Content = "[UPDATED] All privius updates"
            $allPriviusUpdatesL.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})
            
            $temp = $allUpdatesT.Text
            
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
            $item | Add-Member -type NoteProperty -Name 'Status' -Value $global:LoadedTicket.Status
            $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $global:LoadedTicket.deadLine
            $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value $global:LoadedTicket.ticketOwner

            $item | Add-Member -type NoteProperty -Name 'id' -Value $global:LoadedTicket.id
            $item | Add-Member -type NoteProperty -Name 'visible' -Value $global:LoadedTicket.visible

            $global:loadedticket =  $item

            saveChanges
        }      
    })
    
    $solvedB.Add_Click({ 
        addComment
        $global:LoadedTicket.Update = $Global:comment        
        $global:LoadedTicket.status = $statusCB.SelectionBoxItem
        $global:loadedticket.Update += $updateT.Text + "`r`n-------------------------------- Solved ----------------------------`r`n"
        $global:loadedticket.Update += $updateT.Text + "`r`n--------------------------------------------------------------------`r`n"        
        saveChanges
        solvedTicket
        searchForTickets
        $Window.Hide()
    })

    $notSolvableB.Add_Click({
        addComment
        $global:LoadedTicket.Update = $Global:comment
        $global:LoadedTicket.status = $statusCB.SelectionBoxItem
        $global:loadedticket.Update += $updateT.Text + "`r`n----------------------- Not able to be Solved ----------------------`r`n"
        $global:loadedticket.Update += $updateT.Text + "`r`n--------------------------------------------------------------------`r`n"
        saveChanges
        NotsolvedTicket
        searchForTickets
        $Window.Hide()
    })

    $tagTtemp = $tagT.Text

    $tagT.Add_MouseEnter({
        $tagT.Background = [System.Windows.Media.Brushes]::LightGray
        $tagT.Text = "$($tagT.Text) - Left-mouseclick to copy to Clipboard."
    })

    $tagT.Add_MouseLeave({
        $tagT.Background = [System.Windows.Media.Brushes]::Transparent
        $tagT.Text = $tagTtemp
    })
 
    $tagT.Add_MouseLeftButtonDown({
        Set-Clipboard -Value $global:LoadedTicket.Tag
        $tagT.Text = "$tagTtemp - Kopierad!"
    })

    $statusCB.Add_SelectionChanged({ 
        $selectedItem = $args[1].AddedItems[0].Content

        if ( $selectedItem -eq "Not set" ) {
            
            $statusT.Visibility = "Visible"
            
            if ( !($desiredStatus -eq "") ) { 
                $statusT.Text = $desiredStatus
            }
        } else {
            $statusT.Visibility = "Hidden"
        }
    })

    $deadlineB.Add_Click({
        calender
        $DeadlineT.Text = "Deadline: $($Global:dateTime)"

    })

    $resetDeadlineB.Add_Click({
        $DeadlineT.Text = "Deadline: Not set"
        $Global:dateTime = ""
    })

    $visibleB.Add_Click({
        
        if ( $visibleB.Content -eq "Public" ) {
            $visibleB.Content = "Private"
            $visibleT.Text = "Visible: Public"

        } else {
            $visibleB.Content = "Public"
            $visibleT.Text = "Visible: Private"
        }

    })

    [Void]$Window.ShowDialog();
}

function calender () {

    # Written mostly from Copilot with some minor tweaks from me.

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Calender" Height="300" Width="400">
    <Grid>
        <Button x:Name="PrevMonthBtn" Height="30" Content="Backward" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="10,10,0,0" Width="100"/>
        <Button x:Name="NextMonthBtn" Height="30" Content="Forward" VerticalAlignment="Top" HorizontalAlignment="Right" Margin="0,10,10,0" Width="100"/>
        <TextBox x:Name="MonthDisplay" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10,0,0" Width="120" IsReadOnly="True"/>
        <WrapPanel x:Name="DateGrid" Margin="10,50,10,50"/>

        <!-- Skapa en horisontell layout för SelectedDate och ExtraTextBox -->
        <StackPanel Orientation="Horizontal" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,200,0,0">
            <TextBox BorderBrush="Transparent" Height="20" Width="63" IsReadOnly="True" Text="Valt datum:"/>
            <TextBox x:Name="SelectedDateT" Height="20" Width="80" Margin="5,0,0,0"/>
        </StackPanel>

        <!-- Knappar bredvid varandra -->
        <StackPanel Orientation="Horizontal" Height="20" HorizontalAlignment="Center" Margin="0,210,0,0">
            <Button x:Name="OkButton" Content="OK" Width="80"/>
            <Button x:Name="CloseButton" Content="Close" Width="80"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $prevBtn = $window.FindName("PrevMonthBtn")
    $nextBtn = $window.FindName("NextMonthBtn")
    $monthDisplay = $window.FindName("MonthDisplay")
    $dateGrid = $window.FindName("DateGrid")
    $selectedDateT = $window.FindName("SelectedDateT")
    $okBtn = $window.FindName("OkButton")
    $closeBtn = $window.FindName("CloseButton")
    $extraBtn = $window.FindName("ExtraButton")

    $global:currentMonth = (Get-Date).Month
    $global:currentYear = (Get-Date).Year

    function updateCalendar {
        $dateGrid.Children.Clear()
        #$monthDisplay.Text = (Get-Culture).DateTimeFormat.MonthNames[$global:currentMonth - 1] + " " + $global:currentYear
        #$temp = (Get-Culture).DateTimeFormat.MonthNames[$global:currentMonth - 1] + " " + $global:currentYear
        #$monthDisplay.Text = ($temp).ToString("MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))

        $culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
        $date = Get-Date -Year $global:currentYear -Month $global:currentMonth -Day 1
        $monthDisplay.Text = $date.ToString("MMMM yyyy", $culture)

        $daysInMonth = [DateTime]::DaysInMonth($global:currentYear, $global:currentMonth)

        for ($i = 1; $i -le $daysInMonth; $i++) { 
            
            if ( $i -lt 10 ) {
                
                [String]$i = "0"+$i
            }
            
            $txtBox = New-Object System.Windows.Controls.TextBox
            $txtBox.Text = "$i"
            $txtBox.Width = 35
            $txtBox.Height = 35
            $txtBox.IsReadOnly = $true

            $currentDate = "$i $($monthDisplay.Text)"
            
            if ( ((Get-Date).ToString("dd MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))) -eq $currentDate ) {
                $txtBox.Background = [System.Windows.Media.Brushes]::Lightblue
            }

            $txtBox.Add_MouseEnter({
               param($sender, $e)
               if ( !($sender.Background -eq "Gray") ) {
                   $sender.Background = [System.Windows.Media.Brushes]::LightGray
               }
            })

            $txtBox.Add_MouseLeave({
                param($sender, $e)
                if ( !($sender.Background -eq "Gray") ) {
                   $sender.Background = [System.Windows.Media.Brushes]::Transparent

                   if ( ((Get-Date).ToString("dd MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))) -like "*$($sender.Text) $($monthDisplay.Text)*" ) {    
                        $sender.Background = [System.Windows.Media.Brushes]::Lightblue
                   }
                }
            })
            $txtBox.Add_PreviewMouseLeftButtonDown({
                param($sender, $e)

                $sender.Background = [System.Windows.Media.Brushes]::Gray
                $selectedDateT.Text = "$($sender.Text) $([cultureinfo]::GetCultureInfo('en-US').DateTimeFormat.MonthNames[$global:currentMonth - 1]) $currentYear"          
            })

            if ( $i -lt 10 ) {
                
                [int]$i = $i.Substring(1,1)
            }

            $dateGrid.Children.Add($txtBox)
        }
    }

    $prevBtn.Add_Click({
        if ($global:currentMonth -eq 1) {
            $global:currentMonth = 12
            $global:currentYear--
        } else {
            $global:currentMonth--
        }
        updateCalendar
    })

    $nextBtn.Add_Click({
        if ($global:currentMonth -eq 12) {
            $global:currentMonth = 1
            $global:currentYear++
        } else {
            $global:currentMonth++
        }
        updateCalendar
    })

    $okBtn.Add_Click({ 
        $Global:dateTime = $selectedDateT.Text 
        $Window.Hide()
    })
    $closeBtn.Add_Click({ $Window.Hide() })

    updateCalendar

    $window.ShowDialog()
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
                <TextBlock Text="📝 Create New Ticket" FontSize="20" FontWeight="Bold" Margin="0,0,0,20" Foreground="#2C3E50" />

                <TextBlock Text="Issue:" FontSize="16" Foreground="Black" />
                <TextBox Name="issueT"  Height="30" Margin="0,5,0,10" Text="Enter issue description..." Foreground="Black"/>

                <TextBlock Text="Priority:" FontSize="16" Foreground="Black" />
                <ComboBox Name="prioCB" Height="30" Margin="0,5,0,10">
                    <ComboBoxItem Content="Non" IsSelected="True"/>
                    <ComboBoxItem Content="Prio 1"/>
                    <ComboBoxItem Content="Prio 2"/>
                    <ComboBoxItem Content="Prio 3"/>
                </ComboBox>

                <TextBlock Text="Description:" FontSize="16" Foreground="#555555" />
                <TextBox Name="descriptionT"   Height="120" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" Foreground="Black"
                         Margin="0,5,0,10" Text="Enter detailed description..." />
                
                <TextBlock Text="Name:" FontSize="16" Foreground="Black" />
                <TextBox  Name="userT" Height="30" Margin="0,5,0,10" Text="Enter your name..." Foreground="Black"/>

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
                $item | ConvertTo-Json | Out-File -FilePath "$Global:newTickets\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 1" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Global:prio1\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 2" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Global:prio2\$($filtertitle).json"
            } elseif ( $prioCB.SelectionBoxItem -eq "Prio 3" ) {
                $item | ConvertTo-Json | Out-File -FilePath "$Global:prio3\$($filtertitle).json"
            } 
            searchForTickets
            $Window.Hide()
        } else {
        
            Write-Error "You need to fill in every box."
        }
    })

    $closeB.Add_Click({$Window.Hide()})

    $issueT.Add_PreviewMouseDown({ 

        if ( $issueT.Text -eq "Enter issue description..." ) {
            $issueT.Text = ""
        }
    })

    $descriptionT.Add_PreviewMouseDown({ 

        if ( $descriptionT.Text -eq "Enter detailed description..." ) {
            $descriptionT.Text = ""
        }
    })

    $userT.Add_PreviewMouseDown({ 
    
        if ( $userT.Text -eq "Enter your name..." ) {
            $userT.Text = ""
        }
    })

    $Window.Add_PreviewMouseDown({ 
    
        if ( $issueT.Text -eq "" ) {
            $issueT.Text = "Enter issue description..."
        }
        if ( $descriptionT.Text -eq "" ) {
            $descriptionT.Text = "Enter detailed description..."
        }
        if ( $userT.Text -eq "" ) {
            $userT.Text = "Enter your name..."
        }
    })

    [Void]$Window.ShowDialog()    
}

function autoTicket () {
    
    <#
        * Automatic create tickets on specific dates
        * Show a list over all the diffrent autotickets
          - Easyest way to this is in the same way with folders.
    #>
    
$inputXML = @"
<Window x:Class="TicketSystem.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="List over schedule tickets" Height="600" Width="900" UseLayoutRounding="True">
    <Grid>
        <DockPanel LastChildFill="True">

            <!-- Översta filterpanelen -->
            <StackPanel DockPanel.Dock="Top" Background="{DynamicResource {x:Static SystemColors.ActiveCaptionBrushKey}}" Margin="10">
                <WrapPanel Background="{DynamicResource {x:Static SystemColors.ActiveCaptionBrushKey}}">
                    <Button Content="Clear" Name="clearB" Margin="10,0,0,0" Padding="0"
                            Background="#FF975252" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>
                    <Button Content="Filter" Name="filterB" Margin="10,0,0,0" Padding="0"
                            Background="#FF0B1E2B" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>
                    <ComboBox Name="filterCB" Width="130" Height="25" Margin="10,10,0,10">
                        <ComboBoxItem></ComboBoxItem>
                        <ComboBoxItem>January</ComboBoxItem>
                        <ComboBoxItem>February</ComboBoxItem>
                        <ComboBoxItem>March</ComboBoxItem>
                        <ComboBoxItem>April</ComboBoxItem>
                        <ComboBoxItem>May</ComboBoxItem>
                        <ComboBoxItem>June</ComboBoxItem>
                        <ComboBoxItem>July</ComboBoxItem>
                        <ComboBoxItem>August</ComboBoxItem>
                        <ComboBoxItem>September</ComboBoxItem>
                        <ComboBoxItem>October</ComboBoxItem>
                        <ComboBoxItem>November</ComboBoxItem>
                        <ComboBoxItem>December</ComboBoxItem>
                    </ComboBox>
                </WrapPanel>
            </StackPanel>

            <!-- Knappar längst ner -->
            <StackPanel DockPanel.Dock="Bottom" Background="white" Margin="10">
                <WrapPanel>

                    <Button Content="OK" Name="okB" Margin="10,0,0,0" Padding="0"
                            Background="#3498DB" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>

                    <Button Content="New Schedule Ticket" Name="newAutoTicketB" Margin="10,0,0,0" Padding="0"
                            Background="#3498DB" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="140" Height="25" Cursor="Hand"/>

                    <Button Content="Import tickets" Name="importB" Margin="10,0,0,0" Padding="0"
                            Background="#3498DB" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>

                    <Button Content="Delete ticket" Name="deleteB" Margin="10,0,0,0" Padding="0"
                            Background="#FF975252" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>
                </WrapPanel>
            </StackPanel>

            <!-- Lista i mitten -->
            <ListView Name="autoticketListViewT" Foreground="Black" FontSize="15">
                <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                    </Style>
                </ListView.ItemContainerStyle>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="CreateDate" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding createDate}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Assigned to" Width="100">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding AssignedTo}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Ticket Name" Width="300">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ticketName}" TextAlignment="Center" 
                                               ToolTip="{Binding ticketName}" ToolTipService.InitialShowDelay="10" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextAlignment="Center" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Deadline" Width="150">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Deadline}" TextAlignment="Center" />
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
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $clearB = $Window.FindName("clearB")
        $filterB = $Window.FindName("filterB")
        $filterCB = $Window.FindName("filterCB")
        $okB = $Window.FindName("okB")
        $newAutoTicketB = $Window.FindName("newAutoTicketB")
        $deleteB = $Window.FindName("deleteB")
        $importB = $Window.FindName("importB")
        $autoticketListViewT = $Window.FindName("autoticketListViewT")
    }

    catch {
        Write-Warning $_.Exception
        throw
    }

    function updateAutoTicketList () {

        $autoticketListViewT.Items.clear()
        $global:loadedAutotickets.Clear()
        $autoTicketsT = (Get-ChildItem -Path $Global:autoTickets -File).FullName

        if ( !([string]::IsNullOrEmpty($autoTicketsT)) ) {
            $autoTicketsT | ForEach-Object { 

                $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                $json = Get-Content -Path $_ | ConvertFrom-Json

                $tempDate = $json.CreateDate.toString().Substring(0, $json.CreateDate.Length-1)

                $ticket = New-Object PSObject -Property @{
                    ticketName = $temp
                    Status = $json.status
                    Priority = $json.prio
                    ReportedBy = $json.username
                    Date = $json.date
                    AssignedTO = $json.ticketOwner
                    DeadLine = $json.deadLine
                    #createDate = $json.CreateDate
                    createDate = $tempDate
                }

                if ( !([string]::IsNullOrEmpty($filterCB.SelectedItem)) ) {                    if ( !([string]::IsNullOrEmpty($json.CreateDate)) ) { 

                        $tempDate = $json.createDate.split(",").trim()

                        $tempDate | ForEach-Object { 
                            
                            $createDate = $_.replace("-", " ");

                            if ( !([string]::IsNullOrEmpty($createDate)) ) {
                                
                                $culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
                                $parsedDate = [datetime]::ParseExact($createDate, "dd MMMM yyyy", $culture)

                                if ( $filterCB.SelectedItem -like "*$($parsedDate.ToString("MMMM", $culture))*" ) {
                                    [void]$autoticketListViewT.Items.Add($ticket)  
                                    $Global:loadedAutotickets.Add($temp, $_)
                                }
                            }
                                
                        }
                    }
                } else {
                    
                    [void]$autoticketListViewT.Items.Add($ticket)  
                    $Global:loadedAutotickets.Add($temp, $_)
                }
            }
        }
    }
    updateAutoTicketList
    
    $autoticketListViewT.Add_SelectionChanged({
        param($sender, $e)
        $selectedItem = $sender.SelectedItem        
        $global:LastAutoSelectTicket = $selectedItem
        if ( $autoticketListViewT.Items.Count -eq $global:LoadedAutoTickets.count -and !$selectedItem.ticketName -eq "" ) {
            $global:LoadedAutoTicket = Get-Content -Path $global:loadedAutotickets[$selectedItem.ticketName] -ErrorAction SilentlyContinue | ConvertFrom-Json
        } 
    })

    
    $autoticketListViewT.Add_MouseDoubleClick({ createAndUpdateAutoTicket -switch Update })
     
    $newAutoTicketB.Add_Click({createAndUpdateAutoTicket})

    $deleteB.Add_Click({
    
        if ( $autoticketListViewT.SelectedItems.ticketName.Length -gt 0  ) {      
            Remove-Item -Path $global:loadedAutotickets[$global:LastAutoSelectTicket.ticketName]
            updateAutoTicketList
        }
    
    })

    $clearB.Add_Click({ 
        $filterCB.SelectedItem = $null 
        updateAutoTicketList
    })
    
   function configureImport ($importFile, $savedBinding) {

$inputXML = @"
<Window x:Class="WpfApp1.ConfigureImport"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="ConfigureImport" Height="530" Width="600">
    <Grid Margin="40,20,0,20">
        <TextBlock Text="Configure the import" FontSize="20" FontWeight="Bold" Margin="0,0,0,0"/>
        <Grid Margin="20,10,0,10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Orientation="Vertical" HorizontalAlignment="Left" Margin="0,30,0,0">
                <TextBlock Text="Choose which sheet to use" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                <StackPanel Orientation="Horizontal">
                    <ComboBox Name="columnCB" Width="100" Margin="5,5,0,0"/>
                </StackPanel>
            </StackPanel>
            <StackPanel Grid.Row="1" Orientation="Vertical" HorizontalAlignment="Left" Margin="0,30,0,0">
                <TextBlock Text="Choose which column to link to which label" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Title" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="titleCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="Name of ticket" Margin="30,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Discription" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="errorCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="Description of ticket." Margin="30,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Assigned" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="nameCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="Person how is assigned to ticket." Margin="30,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Prio" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="prioCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="Prio 1, Prio 2 or Prio 3" Margin="30,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label Content="CreateDate" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="createDateCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="When it should show up in tickets system" Margin="30,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Deadline" Width="100" Margin="5,7,0,0"/>
                    <ComboBox Name="deadLineCB" Width="100" Margin="30,7,0,0"/>
                    <Label Content="DeadLine" Margin="30,0,0,0"/>
                </StackPanel>
            </StackPanel>

            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,20,0,0">
                <Button Content="OK" Name="okB" Margin="10,0,0,0" Padding="0"
                            Background="#3498DB" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>
                <Button Content="Cancel" Name="cancelB" Margin="10,0,0,0" Padding="0"
                            Background="#FF975252" Foreground="White"
                            FontWeight="Bold" BorderBrush="Transparent"
                            Width="100" Height="25" Cursor="Hand"/>
            </StackPanel>
        </Grid>
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

            $global:columnCB = $Window.FindName("columnCB")
            $global:titleCB = $Window.FindName("titleCB")
            $global:errorCB = $Window.FindName("errorCB")
            $global:nameCB = $Window.FindName("nameCB")
            $global:prioCB = $Window.FindName("prioCB")
            $global:createDateCB = $Window.FindName("createDateCB")
            $global:deadLineCB = $Window.FindName("deadLineCB")
            $okB = $Window.FindName("okB")
            $cancelB = $Window.FindName("cancelB")   
        }

        catch {
            Write-Warning $_.Exception
            throw
        } 

        $sheets = (Get-ExcelSheetInfo -Path $importFile).Name

        $sheets | ForEach-Object { $columnCB.Items.Add($_) }

        if ( $savedBinding ) {
            
            $columns = (Import-Excel -Path $importFile -WorksheetName $savedBinding.sheet)[0]
            $columns = $columns | get-member -MemberType NoteProperty | Select-Object -ExpandProperty Name

            $columns | ForEach-Object {
            
                $titleCB.Items.Add($_)
                $errorCB.Items.Add($_)
                $nameCB.Items.Add($_)
                $prioCB.Items.Add($_)
                $createDateCB.Items.Add($_)
                $deadLineCB.Items.Add($_)
            }    

            $columnCB.SelectedItem = $savedBinding.Sheet
            $titleCB.SelectedItem = $savedBinding.Title
            $errorCB.SelectedItem = $savedBinding.Error
            $nameCB.SelectedItem = $savedBinding.Name
            $prioCB.SelectedItem = $savedBinding.prio
            $createDateCB.SelectedItem = $savedBinding.deadLine
            $deadLineCB.SelectedItem = $savedBinding.createDate
        }

        $columnCB.Add_SelectionChanged({
            
            $columns = (Import-Excel -Path $importFile -WorksheetName $columnCB.SelectedItem)[0]
            $columns = $columns | get-member -MemberType NoteProperty | Select-Object -ExpandProperty Name


            $columns | ForEach-Object {
            
                $titleCB.Items.Add($_)
                $errorCB.Items.Add($_)
                $nameCB.Items.Add($_)
                $prioCB.Items.Add($_)
                $createDateCB.Items.Add($_)
                $deadLineCB.Items.Add($_)
            }            
        })

        $okB.Add_Click({
            $Global:cancelImport = $false
            $Window.Close()   
        })

        $cancelB.Add_Click({
            $Global:cancelImport = $true
            $Window.Close()
        })
        
        [Void]$Window.ShowDialog()
   }

   $importB.Add_Click({ 
        
        $importFile = New-Object windows.forms.openfiledialog   
        $importFile.initialDirectory = "";   
        $importFile.title = "Import tickets from schedual"   
        $importFile.filter = "Schedual|*.xlsx"  #"Schedual|*.xlsx;.csv" 
        [Void]$importFile.ShowDialog()

        $Global:cancelImport = $false

        $data = @()

        if ( $importFile.FileName -like "*xlsx" ) {
            
            # Imports files with Xlsx-format
            if ( !(Get-module -Name ImportExcel) ) {
                Install-Module -Name ImportExcel -Scope CurrentUser
            }

            $ProgressBar = showProgressBar -show $show

            if ( Test-Path -Path "$Global:Path\configureImport.json" ) {
               $saveBindning = Get-Content -Path "$Global:Path\configureImport.json" | ConvertFrom-Json
            }

            if ( $saveBindning.Sheet -or $saveBindning.Title -or $saveBindning.Error -or $saveBindning.Name -or $saveBindning.Prio -or $saveBindning.deadLine -or $saveBindning.createDate ) {
                configureImport -importFile $importFile.FileName -savedBinding $saveBindning
            } else {
                configureImport -importFile $importFile.FileName
            }

            if ( !$cansel ) {
                $item = New-Object PSObject
                $item | Add-Member -type NoteProperty -Name 'Sheet' -Value $columnCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'Title' -Value $global:titleCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'Error' -Value $global:errorCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'Name' -Value $global:nameCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'Prio' -Value $global:prioCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $global:createDateCB.SelectedItem
                $item | Add-Member -type NoteProperty -Name 'createDate' -Value $global:deadLineCB.SelectedItem
              
                $item | ConvertTo-Json | Out-File -FilePath "$Global:Path\configureImport.json"    
            
                $data = Import-Excel -Path $importFile.FileName -WorksheetName $columnCB.SelectedItem
            }
            closeProgressBar -ProgressBar $ProgressBar -Show $show 
        }

        elseif ( $importFile.FileName -like "*csv" ) {
            
            # Imports files with CSV-format
        }

        if ( !$Global:cancelImport ) {
            $data | ForEach-Object {
            
                if ( ![string]::IsNullOrEmpty($_.Moment) ) {
               
                    $filtertitle = $_.Moment.Replace(":", "-")

                    $json = Get-Content -Path "$Global:Path\configureImport.json" | ConvertFrom-Json
                
                    $item = New-Object PSObject
                    $item | Add-Member -type NoteProperty -Name 'Title' -Value $_.($json.title)
                    $item | Add-Member -type NoteProperty -Name 'Computer' -Value ""
                    $item | Add-Member -type NoteProperty -Name 'Tag' -Value $env:COMPUTERNAME 
                    $item | Add-Member -type NoteProperty -Name 'Date' -Value (Get-Date -Format "yymmdd")
                    $item | Add-Member -type NoteProperty -Name 'Error' -Value $_.($json.error)
                    $item | Add-Member -type NoteProperty -Name 'Name' -Value $_.($json.name)
                    $item | Add-Member -type NoteProperty -Name 'Update' -Value ""
                    $item | Add-Member -type NoteProperty -Name 'Username' -Value $env:USERNAME
                    $item | Add-Member -type NoteProperty -Name 'Prio' -Value $_.($json.prio)
                    $item | Add-Member -type NoteProperty -Name 'Status' -Value ""      
                    $item | Add-Member -type NoteProperty -Name 'ID' -Value "$(New-Guid)" 
                    if ( $_.($json.deadline) -is [datetime] ) {
                        $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $_.($json.deadline).ToString("dd-MMMM-yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))
                    }
                    if ( $_.($json.createDate) -is [datetime] ) {
                        $item | Add-Member -type NoteProperty -Name 'createDate' -Value $_.($json.createDate).ToString("dd-MMMM-yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))
                    }
                
                    $item | ConvertTo-Json | Out-File -FilePath "$Global:autoTickets\$($filtertitle).json" -Force      
                }
            }
        }
        updateAutoTicketList
    })

    $filterB.Add_Click({updateAutoTicketList})

    $okB.Add_Click({$Window.hide()})

    [Void]$Window.ShowDialog()
}

function createAndUpdateAutoTicket ($switch) {

$inputXML = @"
<Window x:Class="TicketSystem.NewOrUpdateAutoTicket"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Schedule Ticket" Height="660" Width="720"
        Background="#ECEFF1"
        WindowStartupLocation="CenterScreen">

    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="20">
            <StackPanel>

                <TextBlock Name="headerT" Text="📝 Create a High Priority Automatic Ticket" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,20" Foreground="#37474F"/>

                <TextBlock Text="Issue" FontSize="15" Foreground="#263238" Margin="0,0,0,2"/>
                <TextBox Name="issueT" Height="32" FontSize="14" Padding="6" Margin="0,0,0,12" Background="#FAFAFA" BorderBrush="#B0BEC5"/>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,10" VerticalAlignment="Center">
                    <TextBlock Name="createDateT" Text="Create date:" FontSize="14" Foreground="#455A64" VerticalAlignment="Center"/>
                    <ComboBox Name="createDateCB" Width="120" Margin="10,0,0,0" FontSize="13" Background="#FAFAFA" BorderBrush="#B0BEC5" />
                    <Button Name="createDateB" Content="Add" Width="80" Height="26" FontSize="12" Margin="10,0,0,0" Background="#90CAF9" Foreground="Black" BorderBrush="Transparent"/>
                    <Button Name="deleteDeadlineB" Content="Delete" Width="60" FontSize="12" Height="24" Margin="5,0,0,0" Background="#FFEFDE9A" BorderBrush="Transparent"/>
                    <Button Name="resetCreateDateB" Content="Reset" Width="60" FontSize="12" Height="24" Margin="5,0,0,0" Background="#EF9A9A" BorderBrush="Transparent"/>
                </StackPanel>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Priority:" FontSize="14" Foreground="#455A64" VerticalAlignment="Center"/>
                    <ComboBox Name="prioCB" Margin="10,0,0,0">
                        <ComboBoxItem Content="Prio 1" IsSelected="True"/>
                        <ComboBoxItem Content="Prio 2"/>
                        <ComboBoxItem Content="Prio 3"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Select User:" FontSize="14" Foreground="#455A64" VerticalAlignment="Center"/>
                    <ComboBox Name="selectUserCB" Width="200" Margin="10,0,0,0"/>
                </StackPanel>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,10" VerticalAlignment="Center">
                    <TextBlock Text="Status:" FontSize="15" Foreground="#2E7D32" VerticalAlignment="Center"/>
                    <ComboBox Name="statusCB" Width="160" Margin="10,0,15,0" FontSize="13" Background="#FAFAFA" BorderBrush="#B0BEC5">
                        <ComboBoxItem Content="Not set"/>
                        <ComboBoxItem Content="In progress" IsSelected="True"/>
                        <ComboBoxItem Content="Awaiting response"/>
                        <ComboBoxItem Content="Pending"/>
                        <ComboBoxItem Content="Response"/>
                        <ComboBoxItem Content="External"/>
                        <ComboBoxItem Content="Resolved"/>
                        <ComboBoxItem Content="Reopened"/>
                        <ComboBoxItem Content="Rejected"/>
                        <ComboBoxItem Content="Closed without action"/>
                        <ComboBoxItem Content="Awaiting approval"/>
                        <ComboBoxItem Content="Planned"/>
                    </ComboBox>
                    <TextBox Name="statusT" Width="150" Margin="0,0,0,0" Visibility="Hidden" FontSize="13"
                             Padding="4" Text="Custom text" Foreground="Black" AcceptsReturn="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Background="#FAFAFA" BorderBrush="#B0BEC5"/>
                </StackPanel>

                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Name="deadlineT" Text="Deadline: Not set" FontSize="14" Foreground="#455A64" VerticalAlignment="Center"/>
                    <Button Name="deadlineB" Content="Add" Width="80" FontSize="12" Height="24" Margin="10,0,0,0" Background="#A5D6A7" BorderBrush="Transparent"/>
                    <Button Name="resetDeadlineB" Content="Reset" Width="60" FontSize="12" Height="24" Margin="5,0,0,0" Background="#EF9A9A" BorderBrush="Transparent"/>
                </StackPanel>

                <TextBlock Text="Description" FontSize="15" Foreground="#455A64" Margin="0,4,0,2"/>
                <TextBox Name="descriptionT" Height="120" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"
                         FontSize="13" Background="#FAFAFA" BorderBrush="#B0BEC5" Padding="6" Margin="0,0,0,12"/>

                <TextBlock Text="Your Name" FontSize="15" Foreground="#263238" Margin="0,0,0,2"/>
                <TextBox Name="userT" Height="30" Margin="0,0,0,12" FontSize="14" Background="#FAFAFA" Padding="6" BorderBrush="#B0BEC5"/>

                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,20,0,0">
                    <Button Name="newTicketB" Content="Submit" Width="120" Background="#42A5F5" Foreground="White" Padding="6" FontSize="14" FontWeight="SemiBold"/>
                    <Button Name="closeB" Content="Cancel" Width="100" Background="#B0BEC5" Foreground="White" Padding="6" Margin="10,0,0,0" FontSize="14"/>
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
        $headerT = $Window.FindName("headerT")
        $issueT = $Window.FindName("issueT")
        $newNameT = $Window.FindName("newNameT")
        $renameB = $Window.FindName("renameB")
        $newTicketB = $Window.FindName("newTicketB")
        $closeB = $Window.FindName("closeB")
        $descriptionT = $Window.FindName("descriptionT")
        $ticketNameT = $Window.FindName("ticketNameT")
        $statusCB = $Window.FindName("statusCB")
        $statusT = $Window.FindName("statusT")
        $deadlineT = $Window.FindName("deadlineT")
        $deadlineB = $Window.FindName("deadlineB")
        $resetDeadlineB = $Window.FindName("resetDeadlineB")
        $createDateT = $Window.FindName("createDateT")
        $createDateCB = $Window.FindName("createDateCB")
        $createDateB = $Window.FindName("createDateB")
        $deleteCreateDateB = $Window.FindName("deleteDeadlineB")
        $resetCreateDateB = $Window.FindName("resetCreateDateB")
        $userT = $Window.FindName("userT")
        $selectUserCB = $Window.FindName("selectUserCB")

        $prioCB = $Window.FindName("prioCB")
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    $Global:dateTime = ""
    
    $ticketOwners = Get-Content -Path "$Global:Path\owners.json" | ConvertFrom-Json
    $ticketOwners = $ticketOwners.ticketOwners.split(",").trim()

    $selectUserCB.Items.add("")
    $ticketOwners | ForEach-Object { $selectUserCB.Items.add($_) }

    if ( $switch -eq "Update" ) {

        $headerT.Text = "Update schedule Tickets"
        $newTicketB.Content = "Update"
  
        $issueT.Text = $Global:loadedAutoticket.title
        $descriptionT.Text = $Global:loadedAutoticket.Error
        $desiredStatus = $Global:loadedAutoticket.status
        $userT.Text = $Global:loadedAutoticket.Name
        $selectUserCB.SelectedItem = $Global:loadedAutoticket.ticketOwner

        $notSet = $true
        foreach ($item in $statusCB.Items) {
            if ($item.Content -eq $desiredStatus) {
                $statusCB.SelectedItem = $item
                $notSet = $false
                break
            } else {

                $statusCB.SelectedItem = "Not set"
            }
        }

        if ( $notSet ) {
            foreach ($item in $statusCB.Items) {
            if ($item.Content -eq "Not set") {
                $statusCB.SelectedItem = $item
                break
                } 
            }
            $statusT.Visibility = "Visible"
            $statusT.Text = $desiredStatus
        }
        
        #if ( !($Global:loadedAutoticket.createDate -eq $null -or $Global:loadedAutoticket.createDate -eq "") ) {
        #    $createDateT.Text = "Create Date: $($Global:loadedAutoticket.createDate)"
        #} else {
        #    $createDateT.Text = "Create Date: Not set"
        #}

        $createDates = $Global:loadedAutoticket.createDate.split(",").trim()
        $createDates | ForEach-Object { $temp = $_.replace("-", " "); $createDateCB.Items.add($_) }
 
        if ( !($Global:loadedAutoticket.deadLine -eq $null -or $Global:loadedAutoticket.deadLine -eq "") ) {
            $DeadlineT.Text = "Deadline: $($Global:loadedAutoticket.deadLine)"
        } else {
            $DeadlineT.Text = "Deadline: Not set"
        }
    }

    $newTicketB.Add_Click({
        $filtertitle = $issueT.Text.Replace(":", "-")

        if ( !$switch -eq "update" ) {
            ## New schedual ticket
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
            if ( !($statusCB.SelectionBoxItem -eq "Not set") ) {
                $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusCB.SelectionBoxItem     
            } else {
                $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusT.Text
            } 
            $item | Add-Member -type NoteProperty -Name 'ID' -Value "$(New-Guid)" 

            if ( $DeadlineT.Text -eq "Deadline: Not set" ) {
                $item | Add-Member -type NoteProperty -Name 'deadLine' -Value ""
            } else {
                $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $DeadlineT.Text.replace("Deadline: ", "")
            }
            
            #$item | Add-Member -type NoteProperty -Name 'createDate' -Value $createDateT.Text.replace("Create Date: ", "")
            
            $tempCreateDate = ""
            $createDateCB.Items | ForEach-Object { $temp = $_.replace(" ","-"); $tempCreateDate += "$temp," }
            $item | Add-Member -type NoteProperty -Name 'createDate' -Value $tempCreateDate 


            $item | ConvertTo-Json | Out-File -FilePath "$Global:autoTickets\$($filtertitle).json" -Force

        } else {
            ## Update auro ticket

            if ( !($Global:loadedAutoticket.title -eq $issueT.Text) ) {          
                Rename-Item -LiteralPath $global:loadedAutotickets[$global:LastAutoSelectTicket.ticketName] -NewName "$($issueT.Text).json"
                
            }

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
            $item | Add-Member -type NoteProperty -Name 'Prio' -Value "Prio 1"
            $item | Add-Member -type NoteProperty -Name 'ticketOwner' -Value $selectUserCB.SelectionBoxItem 
            if ( !($statusCB.SelectionBoxItem -eq "Not set") ) {
                $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusCB.SelectionBoxItem     
            } else {
                $item | Add-Member -type NoteProperty -Name 'Status' -Value $statusT.Text
            } 
            $item | Add-Member -type NoteProperty -Name 'ID' -Value $Global:loadedAutoticket.id 
            
            
            if ( $DeadlineT.Text -eq "Deadline: Not set" ) {
                $item | Add-Member -type NoteProperty -Name 'deadLine' -Value ""
            } else {
                $item | Add-Member -type NoteProperty -Name 'deadLine' -Value $DeadlineT.Text.replace("Deadline: ", "")
            }
            
            $tempCreateDate = ""
            $createDateCB.Items | ForEach-Object { $temp = $_.replace(" ","-"); $tempCreateDate += "$temp," }
            $item | Add-Member -type NoteProperty -Name 'createDate' -Value $tempCreateDate 

            $item | Add-Member -type NoteProperty -Name 'createDate' -Value $createDateT.Text.replace("Create Date: ", "")

            $item | ConvertTo-Json | Out-File -FilePath "$Global:autoTickets\$($filtertitle).json" -Force
        }
        
        updateAutoTicketList
        $Window.Hide()
    })

    $statusCB.Add_SelectionChanged({ 
        $selectedItem = $args[1].AddedItems[0].Content

        if ( $selectedItem -eq "Not set" ) {
            
            $statusT.Visibility = "Visible"
            
            if ( !($desiredStatus -eq "") ) { 
                $statusT.Text = $desiredStatus
            }
        } else {
            $statusT.Visibility = "Hidden"
        }
    })

    $createDateB.Add_Click({
        calender
        #$createDateT.Text = "Create Date: $($Global:dateTime)"
        $createDateCB.Items.add($Global:dateTime)
        $createDateCB.SelectedIndex = $createDateCB.Items.Count - 1
        
        $Global:dateTime = ""
    })

    $deadlineB.Add_Click({
        calender
        $deadlineT.Text = "Deadline: $($Global:dateTime)"
        $Global:dateTime = ""
    })


    $resetDeadlineB.Add_Click({
        $deadlineT.Text = "Deadline: Not set"
        $Global:dateTime = ""
    })

    $deleteCreateDateB.Add_Click({

        $createDateCB.Items.Remove($createDateCB.SelectedItem)
    })

    $resetCreateDateB.Add_Click({
        #$createDateT.Text = "Create Date: Not set"
        $createDateCB.items.Clear()
        $Global:dateTime = ""
    })

    $closeB.Add_Click({ $Window.Hide() })

    [Void]$Window.ShowDialog();
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

            $NewT = (Get-ChildItem -Path $Global:newTickets -File).FullName
            $SolvdT = (Get-ChildItem -Path $Global:solvedTickets -File).FullName
            $Global:prio1T = (Get-ChildItem -Path $Global:prio1 -File).FullName
            $Global:prio2T = (Get-ChildItem -Path $Global:prio2 -File).FullName
            $Global:prio3T = (Get-ChildItem -Path $Global:prio3 -File).FullName
            $NotSolvdT = (Get-ChildItem -Path $Global:NotsolvedTickets -File).FullName
            $DeletedT = (Get-ChildItem -Path $Global:deletedTickets -File).FullName

            $temp = $loadedtickets[$global:LastSelectTicket.ticketName].Split("\") | Select-Object -Last 1
            $tempDate = ($temp.Split("_") | Select-Object -Last 1).Replace(".json","")

            $date = Get-Date -Format "yyMMdd"
            
            if ( "$NewT" -notlike "*$($newNameT.Text)*" -or `
                   "$SolvdT" -notlike "*$($newNameT.Text)*" -or ` 
                      "$Global:prio1T" -notlike "*$($newNameT.Text)*" -or `
                          "$Global:prio2T" -notlike "*$($newNameT.Text)*" -or ` 
                             "$Global:prio3T" -notlike "*$($newNameT.Text)*" -or ` 
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
    $item | Add-Member -type NoteProperty -Name 'Prio1R' -Value $Global:prio1R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'Prio2R' -Value $Global:prio2R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'Prio3R' -Value $Global:prio3R.IsChecked
    $item | Add-Member -type NoteProperty -Name 'SolvedR' -Value $SolvedR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'NotSolvedR' -Value $NotSolvedR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'WithOutOwner' -Value $showAllTicketsR.IsChecked
    $item | Add-Member -type NoteProperty -Name 'automove' -Value $Global:autoMove
    $item | Add-Member -type NoteProperty -Name 'user' -Value $Global:ticketOwner
    $item | Add-Member -type NoteProperty -Name 'showWithNoOwners' -Value $showWithNoOwners.IsChecked    
    $item | Add-Member -type NoteProperty -Name 'first' -Value $Global:first
    $item | Add-Member -type NoteProperty -Name 'second' -Value $Global:second
    $item | Add-Member -type NoteProperty -Name 'third' -Value $Global:third

    $item | Add-Member -type NoteProperty -Name 'chooseTicketOwner' -Value $Global:chooseTicketOwner
    

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
        Title="Settings" Height="340" Width="583">
    <Grid>
        <Border Background="White" CornerRadius="10" Padding="10" Margin="10">
            <StackPanel>

                <!-- Användarval -->
                <TextBlock Text="User for ticket system" />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Select User:" FontSize="14" FontWeight="Bold" />
                    <ComboBox Name="selectUserCB" Width="200" Margin="10,0,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="When assigning the ticket you have to choose the person:" FontSize="14" FontWeight="Bold"  />
                    <CheckBox Name="chooseTicketOwnerCB" Width="200" Margin="10,3,0,0" />
                </StackPanel>
                <TextBlock Text="Path to all the tickets to load" />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Path to ticket:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="pathT" TextWrapping="Wrap" Text="TextBox" Width="200" Margin="10,0,0,0"/>
                </StackPanel>
                <TextBlock Text="Move tickets to the correct priority if it`s are below paused as the deadline approaches." />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Move ticket automatic:" FontSize="14" FontWeight="Bold"  />
                    <CheckBox Name="automoveCB" Width="200" Margin="10,3,0,0" />
                </StackPanel>

                <TextBlock Text="Determine when a color should appear in the deadline." />
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="First:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="firstTB" TextWrapping="Wrap" Text="" Width="30" Margin="10,0,10,0" Background="#FFB1E7F3" />
                    <TextBlock Text="Second:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="secondTB" TextWrapping="Wrap" Text="" Width="30" Margin="10,0,10,0"  Background="#FFF0F3B1" />
                    <TextBlock Text="Third:" FontSize="14" FontWeight="Bold" />
                    <TextBox Name="thirdTB" TextWrapping="Wrap" Text="" Width="30" Margin="10,0,0,0"  Background="#FFF3B1B1" />
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
        $chooseTicketOwnerCB = $Window.FindName("chooseTicketOwnerCB")
        $selectUserCB = $Window.FindName("selectUserCB")
        $pathT = $Window.FindName("pathT")
        $saveB = $Window.FindName("saveB")
        $closeB = $Window.FindName("closeB")
        $automoveCB = $Window.FindName("automoveCB")
        $firstTB = $Window.FindName("firstTB")
        $secondTB = $Window.FindName("secondTB")
        $thirdTB = $Window.FindName("thirdTB")       
    }
    catch {
        Write-Warning $_.Exception
        throw
    }

    if ( !(Test-Path -Path "$Global:Path\owners.json") ) {

        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'ticketOwners' -Value 'User1'
        
        $item | ConvertTo-Json | Out-File -FilePath "$Global:Path\owners.json"
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

    $automoveCB.IsChecked = $Global:autoMove

    $chooseTicketOwnerCB.isChecked = $Global:chooseTicketOwner

    $pathT.Text = $Global:Path 

    if (!$Global:Path) {
        $pathT.Text = "A path need to be added"
    }

    $temp = $pathT.Text 
    $pathT.Add_GotFocus({ $pathT.Text = "" })
    $Window.Add_MouseDown({ $pathT.Text = $temp })

    $firstTB.Text = $Global:first
    $secondTB.Text = $Global:second
    $thirdTB.text = $Global:third

    $saveB.Add_Click({
 
        if ( [Int]$firstTB.Text -gt [Int]$secondTB.Text -and [Int]$secondTB.Text -gt [Int]$thirdTB.Text ) {
            $Global:ticketOwner = $selectUserCB.SelectedValue
            $ticketOwnerL.Text = "User: $Global:ticketOwner"
            $temp = ""
            $Global:chooseTicketOwner = $chooseTicketOwnerCB.IsChecked
            $Global:Path = $pathT.Text
            $Global:autoMove = $automoveCB.IsChecked
            $Global:first = $firstTB.Text 
            $Global:second = $secondTB.Text
            $Global:third = $thirdTB.Text
            autosaveSettings
            loadAutosaveSettings
            searchForTickets
            $Window.Hide()
        } else {
            
            Write-Host "The first must be larger than the second and the second must be larger than the third."
        }
    })

    $closeB.Add_Click({$Window.Hide()})
    [Void]$Window.ShowDialog()
}

## TESTING

$Timer1 = New-Object System.Windows.Threading.DispatcherTimer
$Timer1.Interval = [TimeSpan]::FromSeconds(2)

function addColor () {

    if ( !$notsolvedR.IsChecked -or $solvedR.IsChecked ) {

        foreach ($item in $tickets.Items) {
            if ( ![string]::IsNullOrEmpty($item.Deadline) ) {              
                $deadline = Get-Date $item.Deadline -ErrorAction SilentlyContinue
                $limit = $($deadline - (Get-Date)).days
                
                if ( $limit -le $Global:first -and $limit -ge $Global:second ) {
                    $container = $tickets.ItemContainerGenerator.ContainerFromItem($item)
                    if ( $container -ne $null ) {                        
                        $color = [System.Windows.Media.Color]::FromRgb(177, 231, 243) # Blue  
                        $brush = New-Object System.Windows.Media.SolidColorBrush($color)
                        $container.Background = $brush
                    }
                } elseif ( $limit -le $Global:second -and $limit -ge $Global:third ) {
                    $container = $tickets.ItemContainerGenerator.ContainerFromItem($item)
                    if ( $container -ne $null ) { 
                        $color = [System.Windows.Media.Color]::FromRgb(240, 243, 177) # Yellow
                        $brush = New-Object System.Windows.Media.SolidColorBrush($color)
                        $container.Background = $brush
                    }
                } elseif ( $limit -le $Global:third  ) {
                    $container = $tickets.ItemContainerGenerator.ContainerFromItem($item)
                    if ( $container -ne $null ) { 
                        $color = [System.Windows.Media.Color]::FromRgb(243, 177, 177) # Red                     
                        $brush = New-Object System.Windows.Media.SolidColorBrush($color)
                        $container.Background = $brush
                    }
                }
            }
            if ( ![string]::IsNullOrEmpty($item.Recurrent) ) {
                $container = $tickets.ItemContainerGenerator.ContainerFromItem($item)
                if ( $container -ne $null ) {
                    $color = [System.Windows.Media.Color]::FromRgb(252, 230, 211) # Orange 250, 140, 39                      
                    $brush = New-Object System.Windows.Media.SolidColorBrush($color)
                    $container.Background = $brush
                }
            }
        } 
    }
}
$Timer1.add_Tick({addColor})

$Timer1.Start()

$Timer2 = New-Object System.Windows.Threading.DispatcherTimer
$Timer2.Interval = [TimeSpan]::FromSeconds(10)

$Timer2.add_Tick({
    
    # Code from Copilot

    $selectedItem = $tickets.SelectedItem
    $selectedIndex = $tickets.SelectedIndex

    searchForTickets

    if ($selectedItem) {
        $newSelection = $tickets.Items | Where-Object { $_.ticketName -eq $selectedItem.ticketName }

        if ($newSelection) {
            $tickets.SelectedItem = $newSelection
        } elseif ($selectedIndex -ge 0 -and $selectedIndex -lt $tickets.Items.Count) {
            $tickets.SelectedIndex = $selectedIndex
        }
    }

    if ( $Global:autoMove ) {
        $Global:pauseT = (Get-ChildItem -Path $Global:pause -File).FullName
        if ( $Global:pauseT ) {  
            $Global:pauseT | ForEach-Object {
                $temp = ($_.Split("\") | Select-Object -Last 1).ToString().Replace(".json", "")
                $json = Get-Content -Path $_ | ConvertFrom-Json

                if ( ![string]::IsNullOrEmpty($json.deadLine) ) {
                    $deadline = Get-Date $json.deadLine -ErrorAction SilentlyContinue
                    $limit = $($deadline - (Get-Date)).days

                    if ( $limit -le $Global:first -and $limit -ge 0 -or $limit -le 0 ) {
                    
                        if ( $json.prio -eq "Prio 1" ) { 
                           Move-Item -Path $_ -Destination $Global:prio1    
                           searchForTickets 
                        } elseif ( $json.prio -eq "Prio 2" ) {
                           Move-Item -Path $_ -Destination $Global:prio2   
                           searchForTickets 
                        }elseif ( $json.prio -eq "Prio 3" ) {
                           Move-Item -Path $_ -Destination $Global:prio3 
                           searchForTickets 
                        }
                    }
                }
            }
        }
    }
 })

$Timer2.Start()

$Timer3 = New-Object System.Windows.Threading.DispatcherTimer
$Timer3.Interval = [TimeSpan]::FromSeconds(86400).TotalDays

function createTicketOnDate () {

 ## Automatic ticket creation

    $allIDs = @()

    $ticketsToBeCreated = (Get-ChildItem -Path $Global:autoTickets).FullName

    $idCheck = (Get-ChildItem -Path $Global:prio1).FullName

    $idCheck | ForEach-Object { 
        
        if ( ![string]::IsNullOrEmpty($_) ) {
            $json = Get-Content -Path $_ | ConvertFrom-Json
            $allIDs += $json.id 
        }
    }
    if ( ![string]::IsNullOrEmpty($ticketsToBeCreated) ) {
        $ticketsToBeCreated | ForEach-Object { 
            $json = Get-Content -Path $_ | ConvertFrom-Json

            if ( ![string]::IsNullOrEmpty($json.createDate) ) {

                #$createDate = Get-Date $json.createDate -ErrorAction SilentlyContinue
                <#
                $culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
                $date = Get-Date $json.createDate  -ErrorAction SilentlyContinue
                $createDate = $date.ToString("dd MMMM yyyy", $culture)

                $date2 = Get-Date -ErrorAction SilentlyContinue
                $curentDate = $date2.ToString("dd MMMM yyyy", $culture)

                $limit = $($createDate - $curentDate).days
                #>

                $culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

                <#
                 # Before multidate

                $date = [datetime]::ParseExact($json.createDate, "dd MMMM yyyy", $culture)
                $createDate = $date
                $currentDate = Get-Date

                $limit = ($createDate - $currentDate).Days

                if ( $limit -ge 0 -or $limit -le 0 ) {
                
                    if ( !($json.ID -contains $allIDs) ) { 
                        Copy-Item -Path $_ -Destination $Global:prio1  
                    }
                }
                #>

                $tempDate = $json.createDate.split(",").trim()
                $tempDate | ForEach-Object { 
    
                    $createDate = $_.replace("-", " "); 

                    if ( !([string]::IsNullOrEmpty($createDate)) ) {

                        $date = [datetime]::ParseExact($createDate, "dd MMMM yyyy", $culture)
                        $createDate = $date
                        $currentDate = Get-Date
                        
                        $limit = ($createDate - $currentDate).Days
                        
                        if ( $limit -ge 0 -or $limit -le 0 ) {
                
                            if ( !($json.ID -contains $allIDs) ) { 
                            
                                $json.createDate = $createDate.ToString("yyyy-MM-dd")
                                #$json | ConvertTo-Json | out-file -FilePath "$Global:prio1\$($json.title).json" -Force  
                                
                                if ( $json.prio -eq "Prio 1" ) {
                                    $json | ConvertTo-Json | out-file -FilePath "$Global:prio1\$($json.title).json"
                                } elseif ( $json.prio -eq "Prio 2" ) {

                                    $json | ConvertTo-Json | out-file -FilePath "$Global:prio2\$($json.title).json"
                                } elseif ( $json.prio -eq "Prio 3" ) {
                                    $json | ConvertTo-Json | out-file -FilePath "$Global:prio3\$($json.title).json"
                                } 
                            }
                        }
                    }
                }
            }
        }

    searchForTickets 
    }
}
createTicketOnDate

$Timer3.add_Tick({createTicketOnDate})
$Timer3.Start()

$settingsM.add_Click({ settings })
$exitM.add_Click({ $MainWindow.Close() })

$newR.Add_Click({ searchForTickets -show $false })
$solvedR.Add_Click({ searchForTickets -show $true })
$notsolvedR.Add_Click({ searchForTickets -show $true })
$Global:prio1R.Add_Click({ searchForTickets -show $true })
$Global:prio2R.Add_Click({ searchForTickets -show $true })
$Global:prio3R.Add_Click({ searchForTickets -show $true })
$Global:pauseR.Add_Click({ searchForTickets -show $true })
$showWithNoOwners.Add_Click({ searchForTickets })

$Global:noOwnersWasChecked = $false  
$showAllTicketsR.Add_Click({ 
    searchForTickets 
    <#
    if ( $Global:noOwnersWasChecked -and !$showWithNoOwners.IsChecked ) {
        $Global:noOwnersWasChecked = $false
        $showWithNoOwners.IsChecked = $true
    } else {
        $Global:noOwnersWasChecked = $showWithNoOwners.IsChecked
        $showWithNoOwners.IsChecked = $false
    }
    #>
})  

$newTicketB.Add_Click({ newTicket })
$renameTicketB.Add_Click({ renameTicket })
$assignticketB.Add_Click({ assignTicketOwner })
$resetAssignTicketB.Add_Click({ resetTicketOwner })
$choosePrio1B.Add_Click({ prio1 })
$ChoosePrio2B.Add_Click({ prio2 })
$ChoosePrio3B.Add_Click({ prio3 })
$Global:pauseTicketB.Add_Click({ pause })
$deleteTicketB.Add_Click({ deleteTicket })
$searchB.Add_Click({ searchForTickets })

$autoTicketB.Add_Click({ autoTicket })

$searchTB.Add_KeyDown({  
    if ( $_.Key -eq "Enter" ) {
        
        searchForTickets
    }
})

$cleareB.Add_Click({ 
    $searchTB.Text = ""
    searchForTickets
})



$tickets.Add_SelectionChanged({
    param($sender, $e)
    $selectedItem = $sender.SelectedItem
    $global:LastSelectTicket = $selectedItem
    if ( $tickets.Items.Count -eq $loadedtickets.Count -and !$selectedItem.ticketName -eq "" ) {
        $global:LoadedTicket = Get-Content -Path $loadedtickets[$selectedItem.ticketName] -ErrorAction SilentlyContinue | ConvertFrom-Json
    } 
})



$tickets.Add_MouseDoubleClick({ openTicket })

$MainWindow.Add_Closing({
    $Timer1 = $null
    $Timer2 = $null
    autosaveSettings
})

$MainWindow.Add_MouseLeftButtonDown({ $tickets.UnselectAll() })

[Void]$MainWindow.ShowDialog()