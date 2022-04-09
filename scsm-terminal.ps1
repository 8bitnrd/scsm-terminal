$scsmEndPoint = "scsm"
$ciresonWebPotal = "ciresonportal.lab.lcl"




<#########################################################################################################>
<##############################             scirpt starts here             ###############################>                                                                                
<#########################################################################################################>

#loading functions
Function ChopString ($stringToChop, $StringLength)
{

    $choppedString = ""
    If($stringToChop.Length -le $StringLength)
    {
        $missingCharsCount = $StringLength - $stringToChop.Length
        $choppedString = $stringToChop
        While($choppedString.Length -lt $StringLength)
        {   
            $choppedString = $choppedString + " "
        }
    }
    else 
    {
      $choppedString = $stringToChop.Substring(0,$StringLength)  
    }

    return $choppedString
}

$userClass = Get-SCSMClass -Name "System.Domain.User$" -ComputerName $scsmEndPoint
$currentuser = Get-SCSMObject -Class $userClass -Filter "UserName -eq $env:username" -ComputerName $scsmEndPoint

#start user interface
do{
    #get assigned tickets
    $relatedAssignedObj = Get-SCSMRelationshipObject -ByTarget $currentuser -ComputerName $scsmEndPoint | Where-Object {$_.RelationshipID -eq "15e577a3-6bf9-6713-4eac-ba5a5b7c4722"}
 
    #write to console
    Clear-Host
    Write-Host "---------------------------------------------------------------------------------------------------------------"
    Write-Host "|                                            ITEMS ASSIGNED TO ME!                                             |"
    Write-Host "---------------------------------------------------------------------------------------------------------------"
    Write-Host "| Ticket Number |  Title                                                                                       |"
    Write-Host "---------------------------------------------------------------------------------------------------------------"
    #               13 char         20 chars               68 cahrs
    foreach($workItem in $relatedAssignedObj.SourceObject)
    {
        
        $outTicketNumber = ChopString -stringToChop $workItem.Name -StringLength 13 
        $outDescription = ChopString -stringToChop $workItem.DisplayName.Split(":")[1] -StringLength 92

        #$outAffectedUser = ChopString -stringToChop  -StringLength 20
        #$outDescription = ChopString -stringToChop $workItem.Description -StringLength 68 

        If($workItem.ClassName -eq "System.WorkItem.Incident")
        {
            Write-host "| " -ForegroundColor White -NoNewline
            Write-host $outTicketNumber -ForegroundColor Yellow -NoNewline
            Write-host " | "  -ForegroundColor White -NoNewline
            #Write-host $outAffectedUser -ForegroundColor Yellow -NoNewline
            #Write-host " | "  -ForegroundColor White -NoNewline
            Write-host $outDescription  -ForegroundColor Yellow -NoNewline
            Write-host " |" -ForegroundColor White
        }
        elseif($workItem.ClassName -eq "System.WorkItem.ServiceRequest")
        {
            Write-host "| " -ForegroundColor White -NoNewline
            Write-host $outTicketNumber -ForegroundColor Green -NoNewline
            Write-host " | "  -ForegroundColor White -NoNewline
            #Write-host $outAffectedUser -ForegroundColor Green -NoNewline
            #Write-host " | "  -ForegroundColor White -NoNewline
            Write-host $outDescription  -ForegroundColor Green -NoNewline
            Write-host " |" -ForegroundColor White
        }
        elseif($workItem.ClassName -eq "System.WorkItem.ChangeRequest")
        {
            Write-host "| " -ForegroundColor White -NoNewline
            Write-host $outTicketNumber -ForegroundColor Blue -NoNewline
            Write-host " | "  -ForegroundColor White -NoNewline
            #Write-host $outAffectedUser -ForegroundColor Blue -NoNewline
            #Write-host " | "  -ForegroundColor White -NoNewline
            Write-host $outDescription  -ForegroundColor Blue -NoNewline
            Write-host " |" -ForegroundColor White
        }
        else
        {
           # Write-host ("| " + $outTicketNumber + " | " + $outAffectedUser + " | " + $outDescription + " |")
        }
    }

    [System.Console]::WriteLine("---------------------------------------------------------------------------------------------------------------")
    [System.Console]::Write("Press Enter To Refresh or Type Ticket Num: ")
    $selection = Read-Host
    if( $relatedAssignedObj.SourceObject -match $selection -and $selection.Length -gt 0)
    {
        Clear-Host
        $relObjSelected = $relatedAssignedObj.SourceObject  | ? {$_.Name -eq $selection.ToUpper()}
        $scsmObject = Get-SCSMObject -Id $relObjSelected.Id -ComputerName $scsmEndPoint
        $affectedUser = (Get-SCSMRelationshipObject -BySource $scsmObject -ComputerName $scsmEndPoint -Filter "RelationshipId -eq 'dff9be66-38b0-b6d6-6144-a412a3ebd4ce'").TargetObject


        $affectedUserDisplay = ChopString -stringToChop $affectedUser.DisplayName -StringLength 24
        $title = chopString -stringToChop $scsmObject.Title -StringLength 105

        Write-Host "---------------------------------------------------------------------------------------------------------------"
        Write-Host "| Affected User: " -NoNewline
        Write-Host  $affectedUserDisplay -BackgroundColor White -ForegroundColor Black -NoNewline
        Write-Host " Alternate Contact:                                                  |"
        Write-Host "|                                                                                                             |"
        Write-Host "| Title:                                                                                                      |"
        Write-Host ("|   $title" + " |")
        Write-Host "|                                                                                                             |"
        $desctiptionString = $scsmObject.Description.Replace([System.Environment]::NewLine,'')
        $totalLength = $desctiptionString.Length
        $counterOne = 0
        $counterTwo = 106
        $outDescription = "| Dexcription:                                                                                                |" 
        while($counterTwo -lt $totalLength)
        {
                 $lineToAdd = $($desctiptionString[$counterOne..$counterTwo] -join '')
                $outDescription = $outDescription + [System.Environment]::NewLine + "| " +  $lineToAdd  + " |"    
                $counterOne = $counterOne + 106
                $counterTwo = $counterTwo + 106
        }

        $lineToAdd = $($desctiptionString[$counterOne..$counterTwo] -join '')
        $WhiteSpacesNeeded = 106 - $lineToAdd.Length
        $wcounter = 0
        While($wcounter -le $WhiteSpacesNeeded)
        {
            $wSpace =  $wSpace + " "
            $wcounter++
        }
        $outDescription = $outDescription + [System.Environment]::NewLine + '| ' + ( $lineToAdd + $wSpace) + ' |' 
        Write-Host $outDescription
        Write-Host "|                                                                                                             |"                                                   
        Write-Host "---------------------------------------------------------------------------------------------------------------"
        Write-Host "| 1 = Add Comment | 2 = Open In Web Portal  | 3 = Exit                                                        |"
        Write-Host "---------------------------------------------------------------------------------------------------------------"
        Write-Host "Make A Selection: " -NoNewline
        $selection = Read-Host
        if($selection -eq 1)
        {
            if($scsmObject.ClassName -eq "System.WorkItem.ChangeRequest")
            {
                Write-Host "     Comments on CRs are not supported, Press ENTER to open CR     " -BackgroundColor Blue
                Read-Host
                & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -app=https://$ciresonWebPotal/changerequest/edit/$($scsmObject.name)
            }
            else 
            {
           
                Write-Host "---------------------------------------------------------------------------------------------------------------"
                Write-Host "------------------------- Write Your Comment Bellow And Press Enter To Submit Comment -------------------------" 
                Write-Host "---------------------------------------------------------------------------------------------------------------"
                $UserCommentComment = Read-Host

                $UserCommentUser = "$env:username"
                $scsmClass = Get-SCSMClass -Name $($scsmObject.ClassName + "$") -ComputerName $scsmEndPoint
                $GUID = [guid]::NewGuid().ToString()

                $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
                $propDescriptionComment = "Comment"

                switch ($scsmClass.Name)
                {
                    "System.WorkItem.Incident" {$CommentClassName = "AnalystComments"}
                    "System.WorkItem.ServiceRequest" {$CommentClassName = "AnalystCommentLog"}
                    "System.WorkItem.Problem" {$CommentClassName = "Comment"}
                    "System.WorkItem.ChangeRequest" {$CommentClassName = "AnalystComments"}   
                }

                # Create the object projection with properties
                $Projection = @{__CLASS = "$($scsmClass.Name)";
                    __SEED = $scsmObject;
                    $CommentClassName = @{__CLASS = $CommentClass;
                                        __OBJECT = @{Id = $GUID;
                                            DisplayName = $GUID;
                                            ActionType = $ActionType;
                                            $propDescriptionComment = $UserCommentComment;
                                            Title = "$($ActionEnum.DisplayName)";
                                            EnteredBy  = $env:username;
                                            EnteredDate = (Get-Date).ToUniversalTime();
                                            IsPrivate = $IsPrivate;
                                        }
                    }
                }

                #create the projection based on the work item class
                switch ($scsmClass.Name)
                {
                    "System.WorkItem.Incident" {New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection$" -Projection $Projection -ComputerName $scsmEndPoint }
                    "System.WorkItem.ServiceRequest" {New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection$" -Projection $Projection -ComputerName $scsmEndPoint}
                    "System.WorkItem.Problem" {New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType$" -Projection $Projection  -ComputerName $scsmEndPoint}
                    #"System.WorkItem.ChangeRequest" {New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel$" -Projection $Projection  -ComputerName $scsmEndPoint}
                }
            }

        }
        elseif($selection -eq 2)
        {
            # 2 was selected
            switch ($scsmObject.ClassName)
            {
                "System.WorkItem.Incident" {& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -app=https://$ciresonWebPotal/incident/edit/$($scsmObject.name)}
                "System.WorkItem.ServiceRequest" {& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -app=https://$ciresonWebPotal/servicerequest/edit/$($scsmObject.name)}
                "System.WorkItem.Problem" {& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -app=https://$ciresonWebPotal/edit/$($scsmObject.name)}
                "System.WorkItem.ChangeRequest" {& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -app=https://$ciresonWebPotal/changerequest/edit/$($scsmObject.name)}
            }
        }
        else
        {
            #nothing was selected
        }

    }
    elseif($selection.Length -eq 0)
    {
        Clear-Host
    }
    else 
    {
        Write-Host "Invalid Select  !!!" -BackgroundColor Red   
        Start-Sleep -Seconds 1
        Clear-Host
    }
}While($true)
