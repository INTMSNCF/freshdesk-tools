Function laps-to-time($text){
    $request = [regex]::Matches($text, $patternHours)
    $hours = $request[0].Groups  | ? { $_.Name -eq 'h' } | Select-Object -ExpandProperty Value
    $request = [regex]::Matches($text, $patternMinuts)
    $minuts = $request[0].Groups  | ? { $_.Name -eq 'm' } | Select-Object -ExpandProperty Value
    $request = [regex]::Matches($text, $patternSeconds)
    $seconds = $request[0].Groups  | ? { $_.Name -eq 's' } | Select-Object -ExpandProperty Value
    return "$(If ($hours) {'{0:00}' -f $hours} Else {'00'}):$(If ($minuts) {'{0:00}' -f $minuts} Else {'00'}):$(If ($seconds) {'{0:00}' -f $seconds} Else {'00'})"
}

Function Lifecycle-To-Excel-Time($initialDirectory)
{
    process {
        try
        {
            [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Title = 'Origen'
            $OpenFileDialog.initialDirectory = $initialDirectory
            $OpenFileDialog.filter = 'Comma-separated values (*.csv)| *.csv'
            $OpenFileDialog.ShowDialog() | Out-Null
            $inputCSV = $OpenFileDialog.filename
            $patternHours = "(?<h>\d+)h"
            $patternMinuts = "(?<m>\d+)m"
            $patternSeconds = "(?<s>\d+)s"
            $tickets = @()
            Import-Csv $inputCSV | Foreach-Object { 
                $obj = New-Object -TypeName psobject
                foreach ($property in $_.PSObject.Properties)
                {
                    If($property.Name -eq 'Ticket id'){
                        $obj | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
                    } Else {
                        $obj | Add-Member -MemberType NoteProperty -Name $property.Name -Value $(laps-to-time $property.Value)
                    }
                } 
                $tickets += $obj
            }
            $tickets | Format-Table
            $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $OpenFileDialog.Title = 'Destino'
            $OpenFileDialog.initialDirectory = $initialDirectory
            $OpenFileDialog.filter = 'Comma-separated values (*.csv)| *.csv'
            $OpenFileDialog.ShowDialog() | Out-Null
            $exportCsv = $OpenFileDialog.filename
            $tickets | Export-Csv -Path $exportCsv -Delimiter ';'
        }
        catch
        {
            write-host -ForegroundColor DarkRed $_.Exception.Message
            read-host “Press ENTER to end...”
        }
    }
}

Lifecycle-To-Excel-Time
