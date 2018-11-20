$inputstr = @"
\`d.T.~Ed/{2D5D526B-021B-415A-96CF-BFFFBC0599B0}.UserInput\`d.T.~Ed/
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
    foreach ($row in ($xml.Userinputs.Userinput | Where {$_.Type -eq "string" -or $_.Type -eq "richtext"}))
    {
        $temp = New-Object -TypeName PSObject
        $temp | Add-Member -MemberType NoteProperty -Name Question -Value $row.Question
        $temp | Add-Member -MemberType NoteProperty -Name Answer -Value $row.Answer

        $data += $temp
    }

    # Add checkbox answers
    foreach ($row in ($xml.Userinputs.Userinput | Where {$_.Type -eq "bool"}))
    {
        If ($row.Answer -eq "true")
        {
             $Answer = "JA"
        }
        If ($row.Answer -eq "false")
         {
             $Answer = "NEJ"
        }
        $temp = New-Object -TypeName PSObject
        $temp | Add-Member -MemberType NoteProperty -Name Question -Value $row.Question
        $temp | Add-Member -MemberType NoteProperty -Name Answer -Value $Answer
        
        $data += $temp
    }

    # Add Query prompt answers
    foreach ($row in ($xml.Userinputs.Userinput | where {$_.Type -eq "System.SupportingItem.PortalControl.InstancePicker"}))
    {
        $temp = New-Object -TypeName PSObject
        $temp | Add-Member -MemberType NoteProperty -Name Question -Value $row.Question
        $answer_string = ($row.Answer).Replace("&lt;","<").Replace("&gt;",">").Replace("&quot;",'"').Replace('&amp;','&').Replace('&apos;',"'")
        $answer = [xml]$answer_string
        $answer_array = @()
        $answer_array = $answer.Values.Value
        $i = 0
        $array_count = ($answer_array | measure).Count
        if ($array_count -gt 1){
            foreach ($obj in $answer_array) { 
                $AnswersTotal += $answer_array[$i].DisplayName + ", "
                $i += 1
            }
        }
        else{
            foreach ($obj in $answer_array) { 
                $AnswersTotal += $answer_array.DisplayName + ", "
                $i += 1
            }
        } 
        
        $AnswersTotal = $AnswersTotal.TrimEnd(", ")
        $temp | Add-Member -MemberType NoteProperty -Name Answer -Value $AnswersTotal
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