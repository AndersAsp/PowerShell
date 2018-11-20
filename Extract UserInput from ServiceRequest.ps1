$inputstr = @"
\`d.T.~Ed/{E5012256-469E-431C-997E-15A2A2ED930E}.UserInput\`d.T.~Ed/
"@

Try
{
    $xml = [xml]$inputstr
    $IsValidXML = $true
}
Catch
{
    $IsValidXML = $false
}

if ($IsValidXML -eq $true)
{
    # Create data-array
    $data = @()

    # Add string-value answers
    foreach ($row in ($xml.Userinputs.Userinput | Where {$_.Type -eq "string"}))
    {
        $temp = New-Object -TypeName PSObject
        $temp | Add-Member -MemberType NoteProperty -Name Question -Value $row.Question
        $temp | Add-Member -MemberType NoteProperty -Name Answer -Value $row.Answer

        $data += $temp
    }

    # Add checkbox answers
    foreach ($row in ($xml.Userinputs.Userinput | Where {$_.Type -ne "string"}))
    {
        $Answer = ""
        Try
        {
            $rowxml = [xml]$row.Answer
            $DisplayNames = @()
            foreach ($DisplayName in $rowxml.Values.Value)
            {
                $DisplayNames += $DisplayName.DisplayName
            }
            $Answer = [System.String]::Join(",",$DisplayNames)
        }
        Catch
        {
            $Answer = ""
        }
        $temp = New-Object -TypeName PSObject
        $temp | Add-Member -MemberType NoteProperty -Name Question -Value $row.Question
        $temp | Add-Member -MemberType NoteProperty -Name Answer -Value $Answer
        
        $data += $temp
    }

    # Format output
    $counter = 0
    $OutputStr = ""
    foreach ($row in $data)
    {
        if ($counter -ne 0)
        {
            $OutputStr += "`n"
        }
        $OutputStr += "Fråga: $($row.Question)`n"
        $OutputStr += "Svar: $($row.Answer)`n"
        $counter++
    }
}
else
{
    $OutputStr = "Invalid XML. Unable to process."
}