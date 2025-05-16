Add-Type -AssemblyName PresentationFramework

$path = Get-Content -Path .\path.txt

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
            
    $item | ConvertTo-Json | Out-File -FilePath "$path\new\$($filtertitle).json"
    
    searchForTickets
    $Window.Hide()
    } else {
        
            Write-Error "Innan du kan skicka måste information fyllas i!"
    }
})

$closeB.Add_Click({$Window.Hide()})

[Void]$Window.ShowDialog()    