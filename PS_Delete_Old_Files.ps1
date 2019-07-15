# 7/15/19
# CREATED BY: Andrew Pimpo

#####INTENDED PURPOSE:
# This program is inteded to work with the "Find Old Files" PS script.
# It will read all paths from the .cvs file and delete these files
# from those paths. This will automate a long, annoying process of
# finding the files in their directories.


#Counters
$count = 0
$err = 0

#Filepath
[string]$textFilePath = $null
[string]$textFileNAME = "PS_DOF_Deletion_List_$(Get-Date -UFormat '%Y-%m-%d').csv"

#Grab list
function get-List(){

    #Create Var
    $list = ""
    
    #Load assembly
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = "c:\\"
    $dialog.Title = "Select the file containing your file list:"
    $dialog.Filter = "CSV Files | *.csv"
    $Show = $dialog.ShowDialog()

    #Receive OK
    if($Show -eq "OK"){
        
        #Set List path
        $list = $dialog.FileName
    }
    else{
        $list = "User cancelled."
    }

    return $list
}
#Set where file will be saved
function set-Save(){
    
    #Save path string
    [string]$path = $null
    
    #Load file browser
	$openDir = new-object System.Windows.Forms.FolderBrowserDialog
	$openDir.RootFolder = "MyComputer"
	$openDir.Description = "Select a Save Location for your summary file"	

	#Show file browser
	$Show = $openDir.ShowDialog()
    
	#Use selected folder\drive as directory
	if($Show -eq "OK")
	{
		$path = $openDir.SelectedPath
    }
    else{
        $path = "Cancel"
    }

    return $path
}
#Read data from excel sheet
function read-List([string]$listPath, [string]$direct){
    
    #Array for writing to file
    $paths = @()
    
    #Create file in $textFilePath
    New-Item -Path $direct -Name $textFileNAME -ItemType "file" -Value "Summary of Deletion:"
    
    #Get path file and send to deletion loop
    $file = Import-Csv $listPath

    foreach ($item in $file){
        
        $count++

        #Create path string
        $p = "$($item.Directory)\$($item.Name)"

        #Delete, write to host, and add to file
        $res = delete-File $p
        Write-Host "File $count`: $p"
        $p = "$p - $res"
                
        $paths += ,$p
    }

    #Write to File
    for($i=0; $i -lt $paths.Length; $i++){
        Add-Content -path $textFilePath -value $paths[$i]
    }
    
    #Record error count
    $good = $count - $err
    $mssg = "`n`n`t$count files were attempted`n`t$good files were deleted`n`t$err files were not deleted because of errors"

    Add-Content -path $textFilePath -value $mssg
    
    #Confirm with User
    Write-Host "$mssg`n`nSummary file in $textFilePath."
}
#YOU WILL BE DELETED
function delete-File($path){

    #Delete
    $res = "DELETED"
    Remove-Item -path $path

    #Test for deletion
    if((Test-Path $path)){
        $res = "***ERROR***"
        $script:err++
    }
    
    #Return result
    return $res
}

##### - BEGIN - #####
$run = "Y"

#WARNING of the results of running this program
$input = [System.Windows.MessageBox]::Show('NOTICE: This program will delete files AND folders. Check your list before running this script.`rOK to continue.`rCancel to quit.','Warning','OKCancel','Error')
switch ($input){
    'OK'{
            $run = "Y"
        }
    'Cancel'{
            $run = "N"
        }
}

while($run -eq "Y"){
    
    #Find file list
    Write-Host "1) Select the file containing your file list (.csv):"
    $listPath = get-List
    Write-Host "- - Path to list: $listPath"

    #Set save path
    Write-Host "`n2) Select save location for your summary file (.txt):"
    $direct = set-Save
    if($direct -eq "Cancel"){
        
        #Clearlist
        $listPath = $null
        Write-Host "User cancelled."
    }
    else{
        $textFilePath = "$direct\$textFileNAME"
        Write-Host "- - Summary file location: $textFilePath."

        #Use file path to open list and Delete files
        read-List $listPath $direct
    }

    $run = "N"

    #Open the path if it exists
    if($listPath -ne $null){
        ii $textFilePath
    }
}

#

#

#

#

#

#

#

#

#

#END