#!/usr/bin/env powershell
#Requires -RunAsAdministrator
function Get-OpenFiles
{
    <#
            .SYNOPSIS
            This function will display open files from a given Sever / PC
            .DESCRIPTION
            This function will use openfiles.exe to get information about open files against a given Sever / PC.
            It will return an DataTable object with all openfiles or if the -filter property is used it will return an filter Dataview object.
            You can also colse opem files if you use the terminate switch
            .EXAMPLE
            Get-OpenFile -ComputerName fileserver1
            .EXAMPLE
            Get-OpenFiles -computername fs2 -Filter Open_File -query \share\ 
            .EXAMPLE
            get-clippboard|Get-OpenFiles  -Filter Accessed_By -query backupadmin
            .EXAMPLE
            Get-OpenFiles  -Filter Accessed_By -query backupadmin -terminate  
            .PARAMETER computername
            The computer name to query. Just one or multiple.
            .PARAMETER Filter
            Here you can specify which poperty is interesting for your result.  
            .PARAMETER query
            This is your search query , you dont have to provide the exact search query the function will alwas use LIKE
            .PARAMETER terminate
            This parameter will trigger an close/discconect against the openfiles which are returned   

    #>

    [CmdletBinding(SupportsShouldProcess,ConfirmImpact = 'Low')]




    Param
    (
        [Parameter(Position = 0,
                Mandatory,
                ValueFromPipeline,
                ValueFromPipelineByPropertyName,
        HelpMessage = 'What computer name would you like to target?')]
       
        [string[]]$computername,

        [Parameter(Position = 1)]
        [ValidateSet('Open_File','Accessed_By')]
        [AllowEmptyString()] 
        [String]$Filter,


        [Parameter(Position = 2)]
        [AllowEmptyString()] 
        [String]$query,

        [Parameter(Position = 3)]
        
        [switch]$Terminate = $false

    )


    begin {
        If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match 'S-1-5-32-544')) 
        {
            Write-Verbose -Message 'You are an administrator'
        }
        else 
        {
            Write-Verbose -Message 'You need Admin rights to execute openfiles.exe'
            break
        }



   
    }


    process {

        Write-Verbose -Message ('Beginning query  {0}' -f $computername)

       

        # openfile execution
        $temp = & "$env:windir\system32\openfiles.exe" /query /s $computername /fo csv /V
    
        #modify the header row , this makes filtering easier , lazy me 
        $temp[0] = '"Hostname","ID","Accessed_By","Type","Locks","Open_Mode","Open_File"'
   
        $csv = $temp | ConvertFrom-Csv
  
    

    
        # Create Datatable
        $dtFileAccess = New-Object -TypeName System.Data.DataTable -ArgumentList ('FileAccess')
        $cols = @('Hostname', 'ID', 'Accessed_By', 'Locks', 'Open_Mode', 'Open_File')

        # Schema (columns)
        foreach ($col in $cols) 
        {
            $null = $dtFileAccess.Columns.Add($col)
        }

        # Fill the Datatable with Values (rows)
        foreach ($c in $csv) 
        {
            $row = $dtFileAccess.NewRow()
            foreach ($col in $cols) 
            {
                $row[$col] = $c.$col
            }
            $null = $dtFileAccess.Rows.Add($row)
        }

        # This will be used for filtering
       
        # DataView rapid filter
        $dvFileAccess	= New-Object -TypeName System.Data.DataView -ArgumentList ($dtFileAccess)
        

        switch ($Filter){
        
            'Accessed_By' 
            {
                $dvFileAccess.RowFilter = 'Accessed_By'+(" LIKE '%{0}%'" -f $query)
            }
        
            'Open_File' 
            {
                $dvFileAccess.RowFilter = ("Open_File LIKE '%{0}%'" -f $query)
            }

            default 
            {
                If ($Terminate -ne $True)
                {
                    return $dtFileAccess
                }
            }
        }
        # Result


        if ( $Terminate -eq $True )
        {
            #-and $dvFileAccess.GetEnumerator() -gt 0){
            $result = [System.Collections.ArrayList]@()
            foreach ($id in $dvFileAccess.GetEnumerator().id)
            {
                $r = & "$env:windir\system32\openfiles.exe" /disconnect /s $computername /id $id
                if ($r -ne $null)
                {
                    $null = $result.Add($r)
                }
            }
            return $result
        }
    
    
        return $dvFileAccess
        
    }
}







