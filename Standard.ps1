[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Function Get-FileName()
{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $initialDirectory,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $type,
         [Parameter(Mandatory=$false, Position=1)]
         [string] $fileName
    )
	
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    if ( $type -eq "Open" ) {
		$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	} else {
		$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
	}
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All Files (*.*)| *.*"
    if ( $fileName -ne $null ) {
        $OpenFileDialog.FileName = $fileName
    }
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function Get-Folder($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
