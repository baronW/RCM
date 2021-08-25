param(
    [string]$TemplatePath,
    [switch]$Silent
)

#Mull this over $DTI = Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select LocalizedDisplayName from SMS_DeploymentType" | %{$_.LocalizedDisplayName} | Sort-Object
#Get-RCMChildFolder -Root -FolderType User_Collection -Recurse

#region RCM Functions
Function Get-RADtree {
    $OrganizationalUnits = Get-ADOrganizationalUnit -Filter *

    $OUObjects = foreach($OU in $OrganizationalUnits){    

        $Split = $OU.DistinguishedName.Split(',')
        #$Split = "OU=Exchange,DC=corp,DC=justice,DC=govt,DC=nz".Split(',')
        $DC = @($Split | %{if($_ -match "DC=(.+)"){$Matches[1]}}) -join "."
    
        $OUS = @($Split | %{if($_ -match "OU=(.*)"){$Matches[1]}})
        $OUJoin = $OUS[$($OUS.count -1)..0] -join "\"

        #$OU.DistinguishedName
        #$OU.Name
        $OUT = New-Object -TypeName psobject
        $OUT | Add-Member -NotePropertyName "FullName" -NotePropertyValue "$DC\$OUJoin"
        $OUT | Add-Member -NotePropertyName "Name" -NotePropertyValue $OU.Name
        $OUT | Add-Member -NotePropertyName "DistinguishedName" -NotePropertyValue $OU.DistinguishedName
        $OUT
    }
    Foreach($OUO in $OUObjects){
        $ParrentFillName = Split-Path $OUO.FullName
        $Parrent = @($OUObjects | ?{$_.FullName -eq $ParrentFillName})[0]
        $OUO | Add-Member -NotePropertyName "Parent" -NotePropertyValue $Parrent
    }
    $OUObjects 
}
Function Get-RCMChildFolder {
    [CmdletBinding(DefaultParameterSetName='Root')]
    param(
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,ParameterSetName='Root')][switch]$Root,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0,ParameterSetName='Object')][psobject]$Object,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0,ParameterSetName='String')][string]$Path,
        [switch]$Recurse,
        [ValidateSet("Application", "Device_Collection", "User_Collection")]$FolderType="Application",
        [int]$Depth=100,
        [int]$CurrentDepth=0
    )
    begin{
        $ReturnLocaltion = Get-Location
        if($RCMisConnected -or $(Connect-RCM)){
            Set-Location $RCMSiteCode
            $Continue = $true
        }
        else{
            $Continue = $false
        }
        switch ($FolderType) {
            "Application" {
                $ObjectType="6000"
                $BaseFolder="Application"
                continue
            }
            "Device_Collection" {
                $ObjectType="5000"
                $BaseFolder="DeviceCollection"
                continue
            }
            "User_Collection" {
                $ObjectType="5001"
                $BaseFolder="UserCollection"
                continue
            }
        }
    }
    process{
        #Write-Host $pscmdlet.ParameterSetName -ForegroundColor Yellow
        #Write-Host $ObjectType -ForegroundColor Yellow
        if($Continue){
            if($pscmdlet.ParameterSetName -eq 'String'){
                $Path = "$Path".TrimEnd("\")
                $FolderName = $($Path -split "\\")[-1]
                $RootFolders = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=$ObjectType And Name='$FolderName'"
                #FullPathCheck
                foreach($F in $RootFolders){
                    $FP = "$($F.Name)"
                    $FD=$F
                    while($FD.ParentContainerNodeID){
                        $FD = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectTypeName='$ObjectTypeName' And ContainerNodeID='$($FD.ParentContainerNodeID)'"
                        $FP = "$($FD.Name)\$FP"
                    }
                    $F | Add-Member -NotePropertyName "Fullname" -NotePropertyValue "$RCMSiteCode\$BaseFolder\$FP" -Force
                }
                $RootFolder = $RootFolders | ?{$_.Fullname -eq $Path}
            
                $ParrentObject = New-Object -TypeName psobject -Property @{name=$($RootFolder.Name);FullName="$Path";Parent="";ContainerNodeID=$($RootFolder.ContainerNodeID)}
                $Children = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=$ObjectType And ParentContainerNodeID='$($ParrentObject.ContainerNodeID)'"
            }
            elseif($pscmdlet.ParameterSetName -eq 'Object'){
                $ParrentObject = $Object
                $Children = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=$ObjectType And ParentContainerNodeID='$($ParrentObject.ContainerNodeID)'"
            }
            else{
                $RootFolder = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=$ObjectType And ParentContainerNodeID=''"
                $ParrentObject = New-Object -TypeName psobject -Property @{name="$BaseFolder";FullName="$RCMSiteCode\$BaseFolder";Parent="";ContainerNodeID=$($_.ContainerNodeID)}
                $Children = $RootFolder | %{New-Object -TypeName psobject -Property @{name=$($_.Name);FullName="$RCMSiteCode\$BaseFolder\$($_.Name)";Parent=$ParrentObject;ContainerNodeID=$($_.ContainerNodeID)}}
            }
            #Write-Host $ParrentObject.FullName -ForegroundColor Cyan
            foreach($Child in $Children){
                $ChildObject = New-Object -TypeName psobject -Property @{name=$($Child.Name);FullName="$($ParrentObject.FullName)\$($Child.Name)";Parent=$ParrentObject;ContainerNodeID=$($Child.ContainerNodeID)}
                $ChildObject
                if($Recurse -and $($CurrentDepth -lt $Depth) -and $Children){
                    #Write-Host "recurse" -ForegroundColor Magenta
                    $NextDepth = $CurrentDepth + 1
                    Get-RCMChildFolder -Recurse -Depth $Depth -CurrentDepth $NextDepth -Object $ChildObject -FolderType $FolderType

                }
                elseif($Recurse -and $($CurrentDepth -ge $Depth)){
                    Write-Host "Cannot recurse depth exceded" -ForegroundColor Magenta
                }
            }
        }
        else{
            Write-Host "Not Connected to SCCM" -ForegroundColor Yellow
        }
    }
    end{
        Set-Location $ReturnLocaltion
    }

}
function get-ordinal {
    param([int]$Number)
    switch ($Number){
        {$_ -match '11$'}{'th';continue}
        {$_ -match '12$'}{'th';continue}
        {$_ -match '13$'}{'th';continue}
        {$_ -match '1$' }{'st';continue}
        {$_ -match '2$' }{'nd';continue}
        {$_ -match '3$' }{'rd';continue}
        default          {'th';continue}
    }
}
function Get-DateString {
    param($Date=$(Get-Date))
    $Date.ToString("dddd 'the' d'$(get-ordinal -Number $Date.ToString('dd'))' 'of' MMMM, yyyy")
}
function Call-MessageBox{
    param(
        [Parameter(Mandatory=$true)][array]$Items,
        [Parameter(Mandatory=$false)][int]$BoxWidth=300,
        [Parameter(Mandatory=$false)][string]$Label
    )

    <#Example
    [array]$ITEMS = @{Text="Attempt to Use Existing App";Tag="$ApplicationName";Type="Button"}
    $ITEMS       += @{Text="Create as $NewName";Tag="$NewName";Type="Button"}
    $ITEMS       += @{Text="Cancel Operation";Tag="Cancel";Type="Button"}
    $ITEMS       += @{Text="Enter Name Manually";Tag="4";Type="TextBox"}
    $ITEMS       += @{Text=@(1..5);Tag="4";Type="Combobox"}
    #>

    $Global:Output_3ae4595be1c84e0db1e08b30491023e9 = ""

    function New-Button{
        param($Item,$Mother)
        $ButtonItem = New-Object System.Windows.Controls.Button
        $ButtonItem.Tag = $Item.Tag
        $ButtonItem.Content = $Item.Text
        $ButtonItem.add_click({
            $Global:Output_3ae4595be1c84e0db1e08b30491023e9 = $this.Tag
            $ButtonWindow.Close()
        })
        $Mother.AddChild($ButtonItem)
    }
    Function New-TextBox{
        param($Item,$Mother)
        $Action = {
            $Global:Output_3ae4595be1c84e0db1e08b30491023e9 = $this.Parent.Children[0].Text
            Write-Verbose -Message "New Textbox Action"
            $ButtonWindow.Close()
        }
        $SPMother = New-Object System.Windows.Controls.StackPanel
        $SPMother.Orientation = "Horizontal"
        $TextBoxItem = New-Object System.Windows.Controls.TextBox
        $TextBoxItem.Width = $BoxWidth - 26
        $TextBoxItem.Text = $Item.text
        $TextBoxItem.add_GotFocus({
            Write-Verbose "add_GotFocus"
            $this.SelectAll()
        })
        $TextBoxItem.add_GotMouseCapture({
            Write-Verbose "GotMouseCapture"
            $this.SelectAll()
        })
        $TextBoxItem.add_Keyup({
            if($_.Key -eq 'RETURN'){
                $Global:Output_3ae4595be1c84e0db1e08b30491023e9 = $this.Parent.Children[0].Text
                Write-Verbose -Message "New Textbox Action"
                Write-Verbose -Message "TextItem: $($TextBoxItem.SelectedText)"
                $ButtonWindow.Close()
            }
        })
        $SPMother.AddChild($TextBoxItem)
        $ButtonItem = New-Object System.Windows.Controls.Button
        $ButtonItem.Content = " + "
        $ButtonItem.Width = "26"
        $ButtonItem.add_click($Action)
        $SPMother.AddChild($ButtonItem)
        $Mother.AddChild($SPMother)
    }
    Function New-ComboBox{
        param($Item,$Mother)
        $Action = {
            Write-Verbose -Message "ComboBox Action"
            $Global:Output_3ae4595be1c84e0db1e08b30491023e9 = $this.SelectedItem
            #$ComboBoxItem.SelectedItem = 1
            $ButtonWindow.Close()
        }
        $SPMother = New-Object System.Windows.Controls.StackPanel
        $SPMother.Orientation = "Horizontal"
        $ComboBoxItem = New-Object System.Windows.Controls.ComboBox
        $ComboBoxItem.Width = $BoxWidth # - 26
        $Item.text | %{$ComboBoxItem.Items.Add($_) | Out-Null}
        $ComboBoxItem.add_SelectionChanged($Action)
        $SPMother.AddChild($ComboBoxItem)
        $Mother.AddChild($SPMother)
    }

    [xml]$ButtonXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Rory's SCCM Importer" WindowStartupLocation = "CenterScreen"
    SizeToContent = "WidthAndHeight" 
    ShowInTaskbar = "True" 
    Background = "White" 
    ResizeMode = "NoResize"
>
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="MainStackPanel" Width="$BoxWidth">
    </StackPanel>
</Window>
"@

    Add-Type -AssemblyName PresentationFramework
    $ButtonReader=(New-Object System.Xml.XmlNodeReader $ButtonXaml)
    $ButtonWindow=[Windows.Markup.XamlReader]::Load($ButtonReader)

    $MainStackPanel  = $ButtonWindow.FindName('MainStackPanel')

    if($Label){
        $LabelItem = New-Object System.Windows.Controls.Label
        $LabelItem.Content = $Label
        $MainStackPanel.AddChild($LabelItem)
    }

    foreach($Item in $Items){
        switch($Item.Type){
            'button'  {New-Button -Item $Item -Mother $MainStackPanel}
            'TextBox' {New-TextBox -Item $Item -Mother $MainStackPanel}
            'ComboBox' {New-ComboBox -Item $Item -Mother $MainStackPanel}
        }
    }

    $async = $ButtonWindow.Dispatcher.InvokeAsync({$ButtonWindow.ShowDialog()})
    $async.Wait() | Out-Null

    $Global:Output_3ae4595be1c84e0db1e08b30491023e9 

}
Function New-MessageBoxObject {
    Param(
        [Parameter(Mandatory=$true)][string]$Text,
        [Parameter(Mandatory=$false)]$Tag,
        [Parameter(Mandatory=$true)][ValidateSet("TextBox", "Button")]$type,
        [Parameter(Mandatory=$false)][string]$ToolTip,
        [Parameter(Mandatory=$false)][scriptblock]$action
    )
    $OUT = New-Object -TypeName psobject
    $OUT | Add-Member -NotePropertyName "Text" -NotePropertyValue $Text
    $OUT | Add-Member -NotePropertyName "type" -NotePropertyValue $type
    if($ToolTip){$OUT | Add-Member -NotePropertyName "ToolTip" -NotePropertyValue $ToolTip}
    if($Tag){$OUT | Add-Member -NotePropertyName "Tag" -NotePropertyValue $Tag}
    if($action){$OUT | Add-Member -NotePropertyName "action" -NotePropertyValue $action}
    $OUT
}
function Add-RCMDistribution {
    param(
        [Parameter(Mandatory=$true)]$APPName,
        [array]$DistributionPointGroups,
        [array]$DistributionPoints
    )
    if($RCMisConnected -or $(Connect-RCM)){
        if (!($Citrix -or $IsCitrix)){$IsCitrix = $false}
        $ReturnLocaltion = Get-Location
        Set-Location $Global:RCMSiteCode

        try{
            if($DistributionPointGroups){$Groups = $DistributionPointGroups | %{Get-CMDistributionPointGroup -Name $_}
                $Groups.name | %{Start-CMContentDistribution -ApplicationName $APPName -DistributionPointGroupName $_}
            }
            if($DistributionPoints){$Points = $DistributionPoints | %{Get-CMDistributionPoint -Name $_}
                $Points | %{if($_.ItemName -match '\[\"Display=[\\]*([^\"]*?)[\\]*\"\]'){$Matches[1]}} | %{Start-CMContentDistribution -ApplicationName $APPName -DistributionPointName $_}
            }
            Write-Host "Done adding Distribution Point" -ForegroundColor Green
            $true
        }
        catch{
            Write-Host "Failed adding Distribution Point" -ForegroundColor Red
            $false
            Write-Error $_
        }
        Set-Location $ReturnLocaltion.Path
    }
}
function get-RCMDistributionPointGroup {
    if($RCMisConnected -or $(Connect-RCM)){
    $ReturnLocaltion = Get-Location
    Set-Location $Global:RCMSiteCode
        get-CMDistributionPointGroup
    }
    Set-Location $ReturnLocaltion.Path
}
function get-RCMDistributionPoint {
    if($RCMisConnected -or $(Connect-RCM)){
    $ReturnLocaltion = Get-Location
    Set-Location $Global:RCMSiteCode
        get-CMDistributionPoint
    }
    Set-Location $ReturnLocaltion.Path
}
function Connect-RCM {
    param($Site,[switch]$Silent)
    if($Site){$Site = $Site.trim('\').trim(':')}
    if((!$RCMisConnected) -or ($Site -and ($RCMSiteCodeRaw -ne $Site))){
        Try{
            Import-Module "$ENV:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1" -Scope Global -ErrorAction Stop # Import the ConfigurationManager.psd1 module
            [array]$drive0 = Get-PSDrive | ?{$_.Provider.name -eq 'CMSite'}
            if ($drive0.count -gt 1){
                if($Silent){Write-Error -Message "BREAK" -ErrorAction Stop}
                $drive = $drive0 | ?{$_.name -eq $Site}
                if(!$drive){
                    $N = 1
                    $drive0 | %{write-host "$N :: $($_.name)";$N++}
                    $Index = Read-Host -Prompt "Select Site"
                    try{$drive = $drive0[$($Index -1)]}
                    catch{Write-Error -Message $_}
                }
            }
            elseif($drive0.count -lt 1){
                if($Silent){Write-Error -Message "BREAK" -ErrorAction Stop}
                $drive = Read-Host -Prompt "No site detected enter manually"
            }
            else{
                $drive = $drive0[0]
            }
            #Set-Location 'AD0:' # Set the current location to be the site code.
            $Global:RCMisConnected = $true
            $Global:ProviderMachineName = $drive.Root
            $Global:RCMSiteCode = "$($drive.name):"
            $Global:RCMSiteCodeRaw = $drive.name

            Write-Host "Connected to $($RCMSiteCode)" -ForegroundColor Green
            $true
        }
        catch{
            $Global:RCMisConnected = $false
            $Global:ProviderMachineName = $null
            $Global:RCMSiteCode = $null
            $Global:RCMSiteCodeRaw = $null
            $false
        }
    }
    else{
        $true
    }
}
function Test-RCMPath {
    param([string]$Path)
    if(($RCMisConnected -or $(Connect-RCM)) -and $Path){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        
        Test-Path $Path

        Set-Location $ReturnLocaltion.Path
    }
    else{
        $false
    }
}
Function ConvertTo-Hashtable {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]$Object,
        [switch]$RemoveNull
    )
    $Out = @{}
    foreach ($P in $Object.psobject.Properties.name){
        if (!$RemoveNull -or ($Object.$P) -or ($Object.$P -eq 0)){
            $Out.$P = $Object.$P
        }
    }
    $Out
} 
Function Get-IncrementedPackageVersion {
    Param([string]$Name,[switch]$silent)
    $REGEX = ([regex]"(?i)^(.*(R|V|\.|.))([\d]+)(.*?)$").Matches($Name)
    if($REGEX.Success){
        $VersionNumber = $REGEX.Groups[3].Value.ToInt32($null) + 1
        $StringFormat = $REGEX.Groups[3].Value -replace ".","0"
        $REGEX.Groups[1].Value + $VersionNumber.ToString($StringFormat) + $REGEX.Groups[4].Value
    }
    elseif($silent){
        "$Name R02"
    }
    else{
        Write-Error -Message "String not recognised as an application name"
    }
}
function Get-RCMApp {
    param(
        [string]$Name,
        [string]$appID,
        [string]$ModelName,
        [switch]$Fast
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $SPLAT = @{}
        if($Name){$SPLAT+=@{Name=$Name}}
        elseif($appID){$SPLAT+=@{appID=$appID}}
        elseif($ModelName){$SPLAT+=@{ModelName=$ModelName}}
        if($Fast){$SPLAT+=@{Fast=$true}}

        Get-CMApplication @SPLAT

        Set-Location $ReturnLocaltion.Path
    }
}
function Get-RCMAppXml {
    Param(
        [Parameter(Mandatory=$true)][string]$Appname,
        [switch]$Serialize
    )
    if($RCMisConnected -or $(Connect-RCM)){
        #$AppName = "Win10 RoryTestApp 10.0"
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $AppObject = Get-CMApplication -Name $Appname

        $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)
        if ($Serialize){[Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($AppXML)}
        else{$AppXML}
        Set-Location $ReturnLocaltion.Path
    }
}
function Get-RCMCollection {
    param(
        [string]$Name,
        [string]$appID,
        [string]$ModelName,
        [switch]$Fast
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $SPLAT = @{}
        if($Name){$SPLAT+=@{Name=$Name}}
        elseif($appID){$SPLAT+=@{appID=$appID}}
        elseif($ModelName){$SPLAT+=@{ModelName=$ModelName}}
        if($Fast){$SPLAT+=@{Fast=$true}}

        Get-CMCollection @SPLAT

        Set-Location $ReturnLocaltion.Path
    }
}
function Get-RCMDeploymentType {
    param(
        [string]$DeploymentTypeName,
        [string]$ApplicationName

    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        if($DeploymentTypeName){
            Get-CMDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName
        }
        else{
            Get-CMDeploymentType -ApplicationName $ApplicationName
        }

        Set-Location $ReturnLocaltion.Path
    }
}
Function get-RcmPackageDetails {
    param($ContentPath,$Installer=$null)

    if($Installer -and (Test-Path $Installer)){
        $InstallerITEM = Get-Item -Path $Installer
    }

    $ReturnLocaltion = Get-Location
    Set-Location $env:TEMP
    $FolderItems = Get-ChildItem -LiteralPath $ContentPath -File
    $FolderItems | Add-Member -NotePropertyName Weight -NotePropertyValue 000000 -Force -PassThru | Add-Member -NotePropertyName UnWeight -NotePropertyValue 000000 -Force -PassThru | Add-Member -NotePropertyName Technology -NotePropertyValue "NA" -Force
    #$FolderItems | Add-Member -NotePropertyName UnWeight -NotePropertyValue 000000 -Force -PassThru | Add-Member -NotePropertyName Technology -NotePropertyValue "NA" -Force
    
    #in an installer is specified that file gets max weight
    if($InstallerITEM){
        $FolderItems | ?{$_.FullName -eq $InstallerITEM.FullName} | %{$_.weight += 8000000}
    }

    $ScriptPriotity = 1000000
    $AppVPriority = 100000
    $MsiPrioriity = 1000

    $ScriptExtentions = @{
        ".bat"=000700
        ".cmd"=000600
        ".vbs"=000400
        ".ps1"=000100
    }
    #Detect Install Script
    #Weights primaraly based on if the script is called install, with a sub weight based on the type of script
    $Scripts = $FolderItems | ?{$ScriptExtentions.Keys -contains $_.Extension}


    Foreach ($S in $Scripts){
        $S.Weight += $ScriptExtentions.$($s.Extension)
        $S.UnWeight += $ScriptExtentions.$($s.Extension)
        $S.Technology = "Script"
        if(($S.name -like "*Install*") -and ($S.name -notlike "*UnInstall*")){$S.Weight+=100000}
        if(($S.name -like "*UnInstall*") -or ($S.name -like "*remove*")){$S.UnWeight+=100000}
    }

    #Weights On Msi if there is more than one msi weights based on the largest msi
    $N = 90
    Foreach ($M in $FolderItems | ?{$_.Extension -eq ".msi"} | Sort-Object -Property length -Descending){
        $M.Weight += $N*$MsiPrioriity
        $M.Technology = "Msi"
        $N--
    }

    Foreach ($A in $($FolderItems | ?{$_.Extension -eq ".appv"})){
        $A.Weight += $AppVPriority
        $A.Technology = "App-V"
    }
    #if($Installer){$MainInstaller = $Installer}

    $FolderItemsSorted = $FolderItems | Sort-Object weight -Descending
    $MainInstaller = ""
    $MainUnInstaller = ""
    $OUT = New-Object -TypeName psobject
    :Tooth foreach ($I in $FolderItemsSorted){
        $Sane = $false
        switch ($I.Technology){
            "Script" {
                $Sane = $true
                $Unsort = $FolderItems | Sort-Object Unweight -Descending |?{$_.name -ne $I.name}|?{$I.Technology -eq "Script"}
                if($Unsort){$MainUnInstaller = $Unsort[0]}
            }
            "Msi" {
                $Sane = $true
            }
            "App-V" {
                [array]$AppVs = $FolderItems | ?{$_.Extension -eq ".appv"}
                if($AppVs.Count -eq 1){
                    $Sane = $true
                }
            }
            default {}
        }
        if($Sane){
            $MainInstaller = $I
            $OUT | Add-Member -MemberType NoteProperty -Name "Technology" -Value $I.Technology

            break Tooth
        }
    }

    $OUT | Add-Member -MemberType NoteProperty -Name "Installer" -Value $MainInstaller
    $OUT | Add-Member -MemberType NoteProperty -Name "InstallScript" -Value @($FolderItems | ?{$ScriptExtentions.Keys -contains $_.Extension} | Sort-Object Weight -Descending)[0]
    $OUT | Add-Member -MemberType NoteProperty -Name "UnInstallScript" -Value @($FolderItems | ?{$ScriptExtentions.Keys -contains $_.Extension} | Sort-Object Unweight -Descending)[0]
    $OUT | Add-Member -MemberType NoteProperty -Name "Msi" -Value @($FolderItems | ?{$_.Extension -eq ".msi"} | Sort-Object Weight -Descending)[0]
    $OUT | Add-Member -MemberType NoteProperty -Name "Mst" -Value @($FolderItems | ?{$_.Extension -eq ".mst"} | Sort-Object Weight -Descending)[0]
    $OUT | Add-Member -MemberType NoteProperty -Name "App-V" -Value @($FolderItems | ?{$_.Extension -eq ".appv"} | Sort-Object Weight -Descending)[0]

    $OUT
    Set-Location $ReturnLocaltion
}
function Zip-Unzip{
    param(
        [switch]$Unzip,
        [string]$ZipPath,
        [string]$FolderPath,
        [ValidateSet("NoCompression", "Fastest", "Optimal")][string]$CompressionLevel="Fastest"
    )
    Add-Type -Assembly System.IO.Compression.FileSystem
    if($Unzip){
        $N = 1
        $FolderPath0 = $FolderPath
        while((Test-Path $FolderPath0) -and (Get-ChildItem -Path $FolderPath0)){
            $N++
            $FolderPath0 = "$($FolderPath)_$N"
        }
        New-Item -Path $FolderPath0 -ItemType Directory -Force
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath,$FolderPath0)
        $FolderPath0
    }
    else{
        $N = 1
        if($ZipPath -match "^(.*)\.([^\.]*)$"){
            $ZipPath0 = $Matches[1]
            $Extention = $Matches[2]
        }

        while(Test-Path "$ZipPath0.$Extention"){
            $N++
            $ZipPath0 = "$($Matches[1])_$N"
        }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($FolderPath,"$ZipPath0.$Extention", [System.IO.Compression.CompressionLevel]::$CompressionLevel, $false)
        "$ZipPath0.$Extention"
    }
}
Function Get-MSIProperty {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$Path,
        [Parameter(ValueFromPipeline=$false)][Array]$Property='*'
    )
    $ReturnLocaltion = Get-Location
    Set-Location $Env:TEMP
    $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    try{
        $Item = Get-Item -Path $Path
    }
    catch{
        Write-Error -Message "Can't Access $Path or not valid MSI"
        break
    }
    try{
        $MSI = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path, 0))
        #$MSI = $windowsInstaller.OpenDatabase($Path, 0)
        $RemoveCopy=$false
    }
    catch{
        $CopyPath = "$Env:TEMP\$([guid]::NewGuid().Guid).msi"
        Write-Verbose "MSI is in use copying to $CopyPath"
        Copy-Item -Path $Path -Destination $CopyPath
        $RemoveCopy=$true
        try{
            $MSI = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($CopyPath, 0))
            #$MSI = $windowsInstaller.OpenDatabase($CopyPath, 0)
        }
        catch{
            Write-Error -Message "MSI is Unreadable"
            Remove-Item -Path $CopyPath
            break
        }
    }

    $Query = "SELECT Property FROM Property"
    $View = $MSI.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSI, ($Query))
    #$VIEW = $MSI.OpenView($Query)
    $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
    #$View.Execute() | Out-Null 
    #$Record = $($View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null))
    #$Record.GetType().InvokeMember("StringData", "GetProperty", $null,$($View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)),1)
    $MSIProperties = while($true){
        try{
            $Record = $($View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null))
            $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        }
        catch{break}
    }
    #$MSIProperties = while($true){try{$View.Fetch($Null).StringData(1)}catch{break}}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($VIEW) | Out-Null

    $PropertiesToFetch = &{foreach ($P in $Property){foreach ($M in $MSIProperties){if($M -like $P){$M}}}} | Sort-Object -Unique
    
    $Out = New-Object -TypeName psobject 
    foreach($P in $PropertiesToFetch){
        $Query = "SELECT Value FROM Property WHERE Property = '$P'"
        $View = $MSI.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSI, ($Query))
        #$VIEW = $MSI.OpenView($Query)
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
        #$View.Execute() | Out-Null  
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        #$Record = $View.Fetch($Null)
        $Out | Add-Member -NotePropertyName $P -NotePropertyValue $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($VIEW) | Out-Null
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MSI) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null
    $Out
    if($RemoveCopy){
        Write-Verbose "Removing $CopyPath"
        for ($i = 1;(($i -lt 20) -and (Test-Path $CopyPath)); $i++){Remove-Item -Path $CopyPath -Force}
    }
    Set-Location $ReturnLocaltion.Path
}
function New-RCMApp {
    param(
        [Parameter(Mandatory=$true)][Alias("Name")]$APPName,
        [Parameter(Mandatory=$false)][string]$IconFile,
        [Parameter(Mandatory=$false)][string]$Publisher,
        [Parameter(Mandatory=$false)][string]$Version,
        [Parameter(Mandatory=$true)]$SCCMFolder,
        [Parameter(Mandatory=$false)]$Comment,
        [Parameter(Mandatory=$false)]$LocalizedName,
        [Parameter(Mandatory=$false)][Alias("AllowTaskSequence")][switch]$AutoInstall 
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        if (!$LocalizedName){$LocalizedName = $APPName}

        try{
            #Icon Shenagins requires setting location to a file location for an easy life
            if($IconFile){
                #Troms out "" 
                if($IconFile -match '"([^"]*)"'){
                    Write-Verbose "Icon Path is dodgy correcting: $IconFile"
                    $IconFile = $Matches[1]
                    Write-Verbose "$IconFile"
                }

                Set-Location $env:USERPROFILE

                if(Test-Path $IconFile){
                    $ICONitem = Get-Item $IconFile
                    if(@('.JPG','.JPEG','.PNG') -contains $ICONitem.Extension){
                    
                    }
                    elseif(($ICONitem.Extension -eq '.exe') -or ($ICONitem.Extension -eq '.ico')){
                        Write-Host "Extracting Icon" -ForegroundColor Yellow
                        Add-Type -AssemblyName System.Drawing
                        $Format = [System.Drawing.Imaging.ImageFormat]::png
                        $TempLocation = "$env:TEMP\$($ICONitem.Name)"
                        Copy-Item $IconFile -Destination $TempLocation -Force
                        $IconData = [System.Drawing.Icon]::ExtractAssociatedIcon($TempLocation)
                        $IconData.ToBitmap().Save("$TempLocation.png",$Format)
                        Copy-Item "$TempLocation.png" -Destination "$IconFile.png" -Force
                        $IconFile = "$IconFile.png"

                        Remove-Item $TempLocation -ErrorAction SilentlyContinue
                        Remove-Item "$TempLocation.png" -ErrorAction SilentlyContinue
                    }
                    else{
                        Write-Host "Invalid Icon" -ForegroundColor Magenta
                        $IconFile = $null
                    }
                
                }
                else{
                    Write-Host "Icon Not found" -ForegroundColor Magenta
                    $IconFile = $null
                }
                Set-Location $RCMSiteCode
            }
            
            $Parameters = @{
                Name=$APPName
                LocalizedName=$LocalizedName
            }
            if ($AutoInstall){$Parameters += @{AutoInstall=$true}}
            if ($Comment){$Parameters += @{Description=$Comment}}
            if ($IconFile){$Parameters += @{IconLocationFile=$IconFile}}
            if ($Version){$Parameters += @{SoftwareVersion=$Version}}
            if ($Publisher){$Parameters += @{Publisher=$Publisher}}

            $app = New-CMApplication @Parameters

            Move-CMObject -FolderPath $SCCMFolder -InputObject $app

            if($AutoInstall){
                $AppObject = Get-CMApplication $AppName
                $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)  #grabs the XML in a eradible form
                $AppXML.AutoInstall = $true
                $AppObject.SDMPackageXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($AppXML)
                $AppObject.Put()
            }

            $app

            write-Host "Done making and moving App"  -ForegroundColor Green
        }
        catch{
            $false
            write-Host "Failed making and moving App"  -ForegroundColor Red
            Write-Error $_
        }
        Set-Location $ReturnLocaltion.Path
    }
}
function New-RCMAppFromTemplate {
    param($template)
    #Gets the name and App instructions from the template
    $ApplicationName = $Template.Name

    $AppSplat = $Template.app | ConvertTo-Hashtable -RemoveNull
    if($AppSplat.IconFile){$AppSplat.IconFile = "$SP\$($AppSplat.IconFile)"}
    $ExistingApp = Get-RCMApp -Name $ApplicationName
    #makes sure that the app is unique
    if($ExistingApp){
        if($RunningGUI){
            $NewName = $ApplicationName
            do{$NewName = Get-IncrementedPackageVersion -Name $NewName}while(Get-RCMApp -Name $NewName)
            [array]$ITEMS = @{Text="Attempt to Use Existing App";Tag="$ApplicationName";Type="Button"}
            $ITEMS       += @{Text="Create as $NewName";Tag="$NewName";Type="Button"}
            $ITEMS       += @{Text="Cancel Operation";Tag="Cancel";Type="Button"}
            $ITEMS       += @{Text="Enter Name Manually";Tag="4";Type="TextBox"}

            $Read = Call-MessageBox -Items $ITEMS -BoxWidth 600 -Label "$ApplicationName Already Exists"
            if($Read -eq $ApplicationName){
                $AppObject = $ExistingApp
            }
            elseif($Read -eq "Cancel"){
                return
            }
            else{
                $ApplicationName = $NewName
                $AppObject = New-RCMApp -APPName $ApplicationName @AppSplat
            }
        }
        else{
            Write-Host "$ApplicationName Already Exists" -ForegroundColor Yellow
            $NewName = $ApplicationName

            do{$NewName = Get-IncrementedPackageVersion -Name $NewName}while(Get-RCMApp -Name $NewName)
            do{
                Write-Host "1: Attempt to Use Existing App" -ForegroundColor Yellow
                Write-Host "2: Create as $NewName" -ForegroundColor Yellow
                Write-Host "3: Cancel Operation" -ForegroundColor Yellow
    
                $Read = Read-Host -Prompt "?:"
            }while(@("1","2","3") -notcontains $Read)
            if($Read -eq "1"){
                #$ApplicationName = $NewName
                $AppObject = $ExistingApp
            }
            elseif($Read -eq "2"){
                $ApplicationName = $NewName
                $AppObject = New-RCMApp -APPName $ApplicationName @AppSplat
            }
            elseif($Read -eq "3"){
                return
            }
            else{Write-Error "How did this happen?"}
        }
    
    }
    else{
        $AppObject = New-RCMApp -APPName $ApplicationName @AppSplat
    }
    $AppObject
}
function New-RCMApplicationDeployment {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$CollectionName,
        [string]$DeployAction='Install',
        [string]$DeployPurpose='Required',
        [string]$UserNotification='DisplaySoftwareCenterOnly',
        [switch]$Passthru,
        [switch]$OverrideServiceWindow,
        [switch]$RebootOutsideServiceWindow,
        [switch]$PreDeploy,
        [switch]$AllowRepairApp,
        [switch]$SendWakeupPacket,
        [switch]$CloseRunningExe
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $DeploymentSplat = @{
            Name=$Name
            CollectionName=$CollectionName
            DeployAction=$DeployAction
            DeployPurpose=$DeployPurpose
            UserNotification=$UserNotification
        }
        #if($Passthru){$DeploymentSplat.Passthru = $true}
        if($OverrideServiceWindow){$DeploymentSplat.OverrideServiceWindow = $true}
        if($RebootOutsideServiceWindow){$DeploymentSplat.RebootOutsideServiceWindow = $true}
        if($PreDeploy){$DeploymentSplat.PreDeploy = $true}
        if($AllowRepairApp){$DeploymentSplat.AllowRepairApp = $true}
        if($SendWakeupPacket){$DeploymentSplat.SendWakeupPacket = $true}
        
        $Deployment = New-CMApplicationDeployment @DeploymentSplat -AvailableDateTime $($(Get-Date).ToUniversalTime().AddMinutes(30))
        if($CloseRunningExe){
            $Deployment.OfferFlags = $Deployment.OfferFlags + 4
            $Deployment.Put()
        }

        Set-Location $ReturnLocaltion.Path
        if($Passthru){$Deployment}
    }
}
Function New-RCMAppTemplate {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0,ParameterSetName='Name')][string]$ApplicationName=$null,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0,ParameterSetName='Object')]$ApplicationObject=$null
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $AppOut = New-Object -TypeName psobject
        $AppOut | Add-Member -NotePropertyName "IconFile" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "SCCMFolder" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "Comment" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "LocalizedName" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "Publisher" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "Version" -NotePropertyValue ""
        $AppOut | Add-Member -NotePropertyName "AllowTaskSequence" -NotePropertyValue ""

        if($PSCmdlet.ParameterSetName -eq "Name"){
            $AppObject0 = Get-CMApplication $ApplicationName
        }
        else{
            $AppObject0 = $ApplicationObject
        }

        foreach($AppObject in $AppObject0){
            #appSection
            $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)
            $Folder = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_ObjectContainerItem where ObjectType='6000' AND InstanceKey is in (Select ModelName from SMS_Application Where LocalizedDisplayName='$($AppObject.LocalizedDisplayName)')"
            $FolderDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=6000 And ContainerNodeID='$($Folder.ContainerNodeID)'"
            $FolderPath = "$($FolderDetails.Name)"
            while($FolderDetails.ParentContainerNodeID){
                $FolderDetails = $FolderDetails=Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectType=6000 And ContainerNodeID='$($FolderDetails.ParentContainerNodeID)'"
                $FolderPath = "$($FolderDetails.Name)\$FolderPath"
            }
            $FolderPath = "$RCMSiteCode\Application\$FolderPath"
            $AppOut.SCCMFolder = $FolderPath
            $AppOut.AllowTaskSequence=$AppXML.AutoInstall
            $AppOut
        }
        Set-Location $ReturnLocaltion
    }
}
Function New-RCMColection {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$LimitingCollectionName,
        [int]$RefreshHours=2,
        [Alias("Folder","SCCMFolder")][string]$FolderPath="",
        [string]$RefreshType='Periodic',
        [string]$Type,
        [string]$GroupName,
        [switch]$Passthru
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        try{
            $CType = if($Type -match 'Device|User'){$Matches[0]}
            $RefreshSheduleObject = New-CMSchedule -RecurInterval Hours -Start (Get-Date) -RecurCount $RefreshHours 
            #if($Type -like '*user*'){$Type = 'User'}
            $CCollection = New-CMCollection -Name $Name -CollectionType $CType -Comment "Made by $env:USERNAME on $(Get-DateString)" -LimitingCollectionName $LimitingCollectionName -RefreshSchedule $RefreshSheduleObject -RefreshType $RefreshType
            Write-Host "Colection $Name made" -ForegroundColor Green 
            $CCollection | Move-CMObject -FolderPath $FolderPath
            Write-Host "Colection Moved to $FolderPath" -ForegroundColor Cyan 
        }
        catch{
            Write-Error -Message $_
        }

        if ($GroupName){
            Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Name -RuleName $GroupName -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemGroupName like ""$env:USERDOMAIN\\$GroupName""" 
        }

        Move-CMObject -FolderPath $FolderPath -InputObject $CCollection
        if ($Passthru){$CCollection}


        Set-Location $ReturnLocaltion.Path
    }
}
function New-RCMCollectionFromTemplate {
    param($Template,$Application,[switch]$Group)
    foreach($C in $Template.Collection){
        
        $collectionSplat = $C | Select-Object -Property * -ExcludeProperty "Deployment","GroupOU","Purpose","GroupName" | ConvertTo-Hashtable -RemoveNull
        if(!$collectionSplat.Name){$collectionSplat.Name = $Template.Name}
        else{$collectionSplat.Name = $collectionSplat.Name -replace "%ApplicaionName%",$Template.Name}
        
        #We don't want to do any group stuff unless there is a chance it will work
        if($C.GroupName -and $C.GroupOU -and $Group){
            #ConfirmName
            try{
                Get-ADGroup -Identity $C.GroupName -ErrorAction SilentlyContinue | Out-Null
                Write-Host "Group already exists" -ForegroundColor Yellow
                $Message = @()
                $Message += @{Text="Use the Existing Group $($C.GroupName)";Tag='TRUE_24e6d0f2-3236-47da-a53b-a342810dad8b';Type="Button"}
                $Message += @{Text="Cancel (Dont Add a Group)";Tag='FALSE_24e6d0f2-3236-47da-a53b-a342810dad8b';Type="Button"}
                $Message += @{Text="Use the name ($(Get-IncrementedPackageVersion $C.GroupName))";Tag="";Type="TextBox"}
                $NewGroupName = $(Call-MessageBox -Items $Message -Label "$($C.GroupName) Already exists") -replace "^Use\ the\ name\ \((?<name>.*)\)$",'${name}'
            }
            catch{
                $NewGroupName = $C.GroupName
            }
            #Create AD group

            if($NewGroupName -eq 'FALSE_24e6d0f2-3236-47da-a53b-a342810dad8b'){
                Write-Host "Not creating a group" -ForegroundColor Yellow
            }
            elseif($NewGroupName -eq 'TRUE_24e6d0f2-3236-47da-a53b-a342810dad8b'){
                Write-Host "Using existing Group" -ForegroundColor Yellow
                $collectionSplat.GroupName = $C.GroupName
            }
            else{
                $collectionSplat.GroupName = $NewGroupName
                try{
                    Get-ADOrganizationalUnit $C.GroupOU | Out-Null
                    New-ADGroup -Path $C.GroupOU -Name $collectionSplat.GroupName -GroupScope Global | Out-Null
                    Write-Host "Group created" -ForegroundColor Cyan 
                }
                catch{
                    Write-Host "Can't Find OU $($C.GroupOU)" -ForegroundColor Yellow
                }
            }
        }
        
        $ExistingColection = Get-RCMCollection -Name $collectionSplat.Name
        if($ExistingColection){
            Write-Host "$($collectionSplat.Name) Collection type Already Exists" -ForegroundColor Yellow
            $NewName = $collectionSplat.Name

            $Message = @("The Collection Already exists")
            $Message += @{Text="Use the Existing Collection $($C.GroupName)";Tag='TRUE_24e6d0f2-3236-47da-a53b-a342810dad8b';Type="Button"}
            $Message += @{Text="Cancel (Dont create a Collection)";Tag='FALSE_24e6d0f2-3236-47da-a53b-a342810dad8b';Type="Button"}
            $Message += @{Text="Use the name ($(Get-IncrementedPackageVersion $collectionSplat.Name))";Tag="";Type="TextBox"}
            $NewCollectionName = $(Call-MessageBox -Items $Message -Label "$($collectionSplat.Name) Already exists") -replace "^Use\ the\ name\ \((?<name>.*)\)$",'${name}'

            if($NewCollectionName -eq 'TRUE_24e6d0f2-3236-47da-a53b-a342810dad8b'){
                $CollectionObject = $ExistingColection
            }
            elseif($NewCollectionName -eq 'TRUE_24e6d0f2-3236-47da-a53b-a342810dad8b'){
                Write-Host "Canceling Collection" -ForegroundColor Yellow
                return
            }
            else{
                $collectionSplat.Name = $NewCollectionName
                $CollectionObject = New-RCMColection @collectionSplat -Passthru
            }

            <#
            do{$NewName = Get-IncrementedPackageVersion -Name $NewName}while(Get-RCMCollection -Name $NewName)
            do{
                Write-Host "1: Attempt to Use Existing Collection" -ForegroundColor Yellow
                Write-Host "2: Create as $NewName" -ForegroundColor Yellow
                Write-Host "3: Cancel Operation" -ForegroundColor Yellow
    
                $Read = Read-Host -Prompt "?:"
            }while(@("1","2","3") -notcontains $Read)
            if($Read -eq "1"){
                #$DtName = $NewName
                $CollectionObject = $ExistingColection
            }
            elseif($Read -eq "2"){
                $collectionSplat.Name = $NewName
                $CollectionObject = New-RCMColection @collectionSplat -Passthru
            }
            elseif($Read -eq "3"){
                return
            }
            else{Write-Error "How did this happen?"}
            #>
        }
        else{
            $CollectionObject = New-RCMColection @collectionSplat -Passthru
        }

        $CollectionObject
    }
}
function New-RCMDeploymentFromTemplate {
    param($Template,$ApplicationName)
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        foreach($C in $Template.Collection){
            if($C.Deployment){
                $CollectionName = if($C.Name){$C.Name}else{$ApplicationName}
                $DeploymentSplat = $C.Deployment | ConvertTo-Hashtable -RemoveNull
                Write-Host $ApplicationName -ForegroundColor Cyan
                Write-Host $CollectionName -ForegroundColor Cyan
                try{$Deployment = New-RCMApplicationDeployment -Passthru -Name $ApplicationName -CollectionName $CollectionName @DeploymentSplat}
                catch{Write-Host "$_" -ForegroundColor Red}
            }
        }
        Set-Location $ReturnLocaltion
    }
}
function New-RCMCollectionTemplate {
    Param($ApplicationName,[switch]$Deployment)
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        #ListOfCollection-DeploymentLinks
        $DeploymentInfo = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName='$ApplicationName'"

        if($Deployment){$AllAppDeploymentObject = Get-CMApplicationDeployment -Name $ApplicationName}

        #This is each collection that the applcaion is linked to , the object isnot the collection object
        Foreach ($DeploymentObject0 in $DeploymentInfo){
            Set-Location -Path $RCMSiteCode
            #this is the collection Object
            $CollectionObject = Get-CMCollection -Name $DeploymentObject0.CollectionName
            $WmiColectionObject = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_Collection Where Name='$($DeploymentObject0.CollectionName)'"

            switch ($WmiColectionObject.CollectionType){
                1 {
                    $CollectionType = "UserCollection"
                    $ObjectTypeName = 'SMS_Collection_User'
                }
                2 {
                    $CollectionType = "DeviceCollection"
                    $ObjectTypeName = 'SMS_Collection_Device'
                }
            }
            

            $Folder = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_ObjectContainerItem where ObjectTypeName='$ObjectTypeName' AND InstanceKey is in (Select CollectionID from SMS_Collection Where Name='$($DeploymentObject0.CollectionName)')"
            $FolderDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectTypeName='$ObjectTypeName' And ContainerNodeID='$($Folder.ContainerNodeID)'"
            $FolderPath = "$($FolderDetails.Name)"
            while($FolderDetails.ParentContainerNodeID){
                $FolderDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "select * from SMS_ObjectContainerNode where ObjectTypeName='$ObjectTypeName' And ContainerNodeID='$($FolderDetails.ParentContainerNodeID)'"
                $FolderPath = "$($FolderDetails.Name)\$FolderPath"
            }

            $out = New-Object -TypeName psobject
            $out | Add-Member -NotePropertyName "Name" -NotePropertyValue ""
            $out | Add-Member -NotePropertyName "LimitingCollectionName" -NotePropertyValue $CollectionObject.LimitToCollectionName
            $out | Add-Member -NotePropertyName "Type" -NotePropertyValue $CollectionType 
            $out | Add-Member -NotePropertyName "Folder" -NotePropertyValue "$RCMSiteCode\$CollectionType\$FolderPath"
            #$out | Add-Member -NotePropertyName "Purpose" -NotePropertyValue $DeploymentObject0.TargetSubName
            $out | Add-Member -NotePropertyName "GroupName" -NotePropertyValue ""
            $out | Add-Member -NotePropertyName "GroupOU" -NotePropertyValue ""
            #AttemptToFindGroupou
            $out.GroupOU = @($CollectionObject.CollectionRules | ?{$_.QueryExpression} | %{if($_.QueryExpression -match "Group[^`"]*`"([^`"]*)`""){try{$(Get-ADGroup $($Matches[1] -replace '.*\\\\','') -ErrorAction SilentlyContinue).DistinguishedName -replace '^cn=[^,]*,',''}catch{}}} | ?{$_})[0]

            IF($Deployment){               

                $AppDeploymentObject = $AllAppDeploymentObject | ?{$_.CollectionName -eq $DeploymentObject0.CollectionName}
                if($AppDeploymentObject){
                    if($AppDeploymentObject.NotifyUser){
                        $UserNotification = "DisplayAll"
                    }
                    elseif($AppDeploymentObject.UserUIExperience){
                        $UserNotification = "DisplaySoftwareCenterOnly"
                    }
                    else{
                        $UserNotification = "HideAll"
                    }

                    #$DeploymentObject = Get-CMDeployment -CollectionName $DeploymentObject0.CollectionName
                    switch ($DeploymentObject0.DeploymentIntent){
                        0 {$DeployPurpose = 'Required'}
                        2 {$DeployPurpose = 'Available'}
                    }
                    switch ($AppDeploymentObject.DesiredConfigType){
                        1 {$DeployAction = 'Install'}
                        2 {$DeployAction = 'Uninstall'}
                    }



                    $DeploymentOut = New-Object -TypeName psobject 
                    $DeploymentOut | Add-Member -NotePropertyName "DeployAction" -NotePropertyValue $DeployAction #$Collection.TargetSubName
                    $DeploymentOut | Add-Member -NotePropertyName "DeployPurpose" -NotePropertyValue $DeployPurpose
                    $DeploymentOut | Add-Member -NotePropertyName "UserNotification" -NotePropertyValue $UserNotification
                    $DeploymentOut | Add-Member -NotePropertyName "OverrideServiceWindow" -NotePropertyValue $AppDeploymentObject.OverrideServiceWindows
                    $DeploymentOut | Add-Member -NotePropertyName "RebootOutsideServiceWindow" -NotePropertyValue $AppDeploymentObject.RebootOutsideOfServiceWindows
                    $DeploymentOut | Add-Member -NotePropertyName "SendWakeupPacket" -NotePropertyValue $AppDeploymentObject.WoLEnabled
                    switch ($AppDeploymentObject.OfferFlags){
                        {$_ -band 1}{$DeploymentOut | Add-Member -NotePropertyName "PreDeploy" -NotePropertyValue $true}
                        {$_ -band 4}{$DeploymentOut | Add-Member -NotePropertyName "CloseRunningExe" -NotePropertyValue $true}
                        {$_ -band 8}{$DeploymentOut | Add-Member -NotePropertyName "AllowRepairApp" -NotePropertyValue $true}
                    }

                    $out | Add-Member -NotePropertyName "Deployment" -NotePropertyValue $DeploymentOut
                }
            }
            
            $out

        }
        Set-Location $ReturnLocaltion
    }
}
Function New-RCMDetectionMethod {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("WindowsInstaller", "Registry", "FileSystem")][string]$Type,
        [Parameter(Mandatory=$false)][ValidateSet("And", "Or")][string]$Connector,
        [Parameter(Mandatory=$true)]$Instructions
    )
    if(!(Connect-RCM -Silent)){return}
    
    switch ($Instructions.Operator){
        'Equals'                   {$ExpressionOperator = 'IsEquals'}
        'Not equal to'             {$ExpressionOperator = 'NotEquals'}
        'Greater than'             {$ExpressionOperator = 'GreaterThan'}
        'Less than'                {$ExpressionOperator = 'LessThan'}
        'Begins with'              {$ExpressionOperator = 'BeginsWith'}
        'Does not begin with'      {$ExpressionOperator = 'NotBeginsWith'}
        'Ends with'                {$ExpressionOperator = 'EndsWith'}
        'Does not end with'        {$ExpressionOperator = 'NotEndsWith'}
        'Contains'                 {$ExpressionOperator = 'Contains'}
        'Does not contain'         {$ExpressionOperator = 'NotContains'}
        'One of'                   {$ExpressionOperator = 'OneOf'}
        'None of'                  {$ExpressionOperator = 'NoneOf'}
        'Between'                  {$ExpressionOperator = 'Between'}
        'Greater than or equal to' {$ExpressionOperator = 'GreaterEquals'}
        'Less than or equal to'    {$ExpressionOperator = 'LessEquals'}
        Default                    {$ExpressionOperator = 'IsEquals'}
    }
    Write-Host $ExpressionOperator -ForegroundColor Cyan
    if($Type -eq "WindowsInstaller" -and $Instructions.ProductCode){
        if($Instructions.ProductCode -match "[\{]{0,1}([0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12})[\}]{0,1}"){
            $GUID = "{$($Matches[1])}"
            if($Instructions.ExpectedValue){
                $DM = New-CMDetectionClauseWindowsInstaller -ProductCode $GUID -Value -ExpectedValue $Instructions.ExpectedValue -ExpressionOperator $ExpressionOperator
            }
            else{
                $DM = New-CMDetectionClauseWindowsInstaller -ProductCode $GUID -Existence
            }
        }
        else{
            Write-Host "ProductCode not valid" -ForegroundColor Red
        }
    }
    elseif($Type -eq 'Registry'){
        $KeyPath = $Instructions.KeyPath -replace "^HKEY_LOCAL_MACHINE\\","" -replace "^HKLM\\","" -replace "^HKLM\:\\","" 
        switch ($Instructions.RegistryHive){
            'HKEY_CLASSES_ROOT'   {$Hive = 'ClassesRoot'}
            'HKEY_CURRENT_USER'   {$Hive = 'CurrentUser'}
            'HKEY_LOCAL_MACHINE'  {$Hive = 'LocalMachine'}
            'HKEY_USERS'          {$Hive = 'Users'}
            'HKEY_CURRENT_CONFIG' {$Hive = 'CurrentConfig'}
            Default {$Hive = 'LocalMachine'}
        }
        
        if($Instructions.RegistryValueName){
            if($Instructions.RegistryValue){

                $DM = New-CMDetectionClauseRegistryKeyValue -Value -Hive $Hive -KeyName $KeyPath -ValueName $Instructions.RegistryValueName -PropertyType $Instructions.RegistryValueType -ExpectedValue $Instructions.RegistryValue -ExpressionOperator $ExpressionOperator -Is64Bit
            }
            else{
                $DM = New-CMDetectionClauseRegistryKeyValue  -Existence -Hive $Hive -KeyName $KeyPath -ValueName $Instructions.RegistryValueName -PropertyType $Instructions.RegistryValueType -Is64Bit
            }
        }
        else{
            $DM = New-CMDetectionClauseRegistryKey -Hive $Hive -KeyName $KeyPath -Is64Bit
        }
    }
    elseif($Type -eq 'FileSystem'){
        switch ($Instructions.Property){
            'Date Modified' {$PropertyType = 'DateModified'}
            'Date Created'  {$PropertyType = 'DateCreated'}
            'Version'       {$PropertyType = 'Version'}
            'Size (Bytes)'  {$PropertyType = 'Size'}
            'Existence'     {$PropertyType = 'Existence'}
            Default         {$PropertyType = 'Version'}
        }
        if($Instructions.Type -eq 'File'){
            if($Instructions.path -match '^\"*(.+)\\([^\\]+?)\"*$'){
                $Path = $Matches[1]
                $File = $Matches[2]
                #Write-Host "$File" -ForegroundColor Cyan
                if($Instructions.Property -eq 'Existence' -or !($Instructions.Value)){
                    $DM = New-CMDetectionClauseFile -FileName $File -Path $Path -Existence -Is64Bit
                }
                else{
                    $DM = New-CMDetectionClauseFile -FileName $File -Path $Path -Value -PropertyType $PropertyType -ExpectedValue $Instructions.Value -ExpressionOperator $ExpressionOperator -Is64Bit
                }
            }
            else{
                Write-Host "PATH not valid" -ForegroundColor Red
            }
        }
        elseif($Instructions.Type -eq 'Folder'){
            if($Instructions.path -match "^(.+)\\([^\\]+)[\\]{0,1}$"){
                $Path = $Matches[1]
                $DirectoryName = $Matches[2]
                if($Instructions.Property -eq 'Existence' -or !($Instructions.Value)){
                    $DM = New-CMDetectionClauseDirectory -DirectoryName $DirectoryName -Path $Path -Existence -Is64Bit
                }
                else{
                    $DM = New-CMDetectionClauseDirectory -DirectoryName $DirectoryName -Path $Path -Value -PropertyType $PropertyType -ExpectedValue $Instructions.Value -ExpressionOperator $ExpressionOperator -Is64Bit
                }
            }
            else{
                Write-Host "PATH not valid" -ForegroundColor Red
            }
        }
    }
    if($DM){
        if($Connector){$DM.Connector = $Connector}
        New-Object -TypeName psobject -Property @{Method=$DM}
    }
    else{
        Write-Host "INVALID INSTRUCTIONS: Creating Placeholder FIX THIS IN THE GUI!" -ForegroundColor Red 
        New-CMDetectionClauseWindowsInstaller -ProductCode "{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}"-Existence
    }
}
function New-RCMMSIDeploymentType {
    param(
        [Parameter(Mandatory=$true)]$APPName,
        [Parameter(Mandatory=$true)]$DeploymentTypeName,
        [Parameter(Mandatory=$true)]$ContentLocation,
        [Parameter(Mandatory=$true)]$MsiName,
        #[Parameter(Mandatory=$false)]$RepairCommand,
        $MsTName,
        $MaximumRuntimeMins=120,
        $EstimatedRuntimeMins=5,
        $UserInteractionMode="Hidden",
        $InstallationBehaviorType='InstallForSystem',
        $SlowNetworkDeploymentMode='Download',
        $LogonRequirementType='WhetherOrNotUserLoggedOn',
        $ContentFallback=$true,
        [array]$DetectionMethods=$null,
        $CloseExecutables
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $MSIPath = "$ContentLocation\$MsiName"
        $MSTPath = "$ContentLocation\$MstName"

        if (!(Test-Path "FileSystem::$MSIPath")){
            Write-Host "MSI path is invalid" -ForegroundColor Yellow
            break
        }

        if ($MSTPath -and !(Test-Path "FileSystem::$ContentLocation\$MstName")){
            Write-Host "MST path is invalid" -ForegroundColor Yellow
            break
        }

        if ($MstName){$InstallCommand = "msiexec /i ""$MsiName"" TRANSFORMS=""$MstName"" /qn"}
        else{$InstallCommand = "msiexec /i ""$MsiName"" /qn"} 

        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        try{

            $MsiProperties = Get-MsiProperty -path "$MSIPath" -Property ProductCode,ProductVersion,ProductName
            $MSIProductCode = $MsiProperties.ProductCode
            Write-Verbose "$MSIProductCode"
            if($MsiProperties.ProductCode){$ProductCode = $MsiProperties.ProductCode}
            else{$ProductCode = '{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}'}
            $commands = @{
                ApplicationName=$APPName
                DeploymentTypeName=$DeploymentTypeName
                ContentLocation=$MSIPath 
                SlowNetworkDeploymentMode=$SlowNetworkDeploymentMode
                MaximumRuntimeMins=$MaximumRuntimeMins
                EstimatedRuntimeMins=$EstimatedRuntimeMins
                Force=$true
                SourceUpdateProductCode=$MSIProductCode
                EnableBranchCache=$true
                ContentFallback=$ContentFallback
                UserInteractionMode=$UserInteractionMode
                InstallCommand=$InstallCommand
                InstallationBehaviorType=$InstallationBehaviorType
                ProductCode=$MsiProperties.ProductCode
            }
            if($MsiProperties.ProductCode){$commands.RepairCommand = "msiexec.exe /fa $($MsiProperties.ProductCode) /qb"}
            if($DetectionMethods){
                Write-Host $DetectionMethods.Count  -ForegroundColor Red
                #if($DetectionMethods){
                    #$commands.AddDetectionClause  = $DetectionMethods[0].Method
                #}Method
                #if($DetectionMethods.Count -gt 1){
                    [array]$AdditionalDetectionMethods = $DetectionMethods | %{$_.Method}
                #}
            }
            else{
                try{
                     $DetectionMethod = New-CMDetectionClauseWindowsInstaller -ExpressionOperator IsEquals -PropertyType ProductVersion -ExpectedValue $MsiProperties.ProductVersion -ProductCode $MsiProperties.ProductCode -Value
                }
                catch{
                    Write-Host "WARNING something went wrong in the detection method CHECK IT" -ForegroundColor Yellow
                    $DetectionMethod = New-CMDetectionClauseWindowsInstaller -ProductCode "{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}" -Existence
                }
                $commands  += @{AddDetectionClause=$DetectionMethod}
            }

            $deploymentType = Add-CMMsiDeploymentType @commands 
            #Add-CMMsiDeploymentType -AddDetectionClause

            #$deploymentType = Get-CMDeploymentType -DeploymentTypeName "Test 8" -ApplicationName "Test app2"

            #this alters some stuff
            $AppObject = Get-CMApplication $AppName
            $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)  #grabs the XML in a eradible form

            $DTXML = $AppXML.DeploymentTypes | ?{"$($_.Scope)/$($_.Name)" -eq $deploymentType.ModelName}
            $DTXML.Installer[0].PostInstallBehavior = 'NoAction' #CHECK

            try{
                #AddsExecutablesToClose
                if($CloseExecutables){
                    foreach($C in $CloseExecutables){
                        $ProcessInfo = [Microsoft.ConfigurationManagement.ApplicationManagement.ProcessInformation]::new()
                        $ProcessInfo.Name = $C
                        $DTXML.Installer[0].InstallProcessDetection.ProcessList.Add($ProcessInfo)
                    }
                }
            }
            catch{
                Write-Host "Unable to add the Close Executables Process List"
            }

            $DetectionLogicalName = @($DTXML.Installer[0].EnhancedDetectionMethod.Settings)[0].LogicalName
            $AppXML.AutoInstall = $true
            $AppObject.SDMPackageXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($AppXML)
            $AppObject.Put()
            #$deploymentType = 
            #foreach($ADT in $AdditionalDetectionMethods){
                #Get-CMDeploymentType -ApplicationName $APPName -DeploymentTypeName $deploymentType.LocalizedDisplayName | Set-CMMsiDeploymentType -AddDetectionClause $ADT.Method
            #    $deploymentType | Set-CMMsiDeploymentType -AddDetectionClause $ADT.Method -RemoveDetectionClause 
            #}
            if($AdditionalDetectionMethods){
                $deploymentType | Set-CMMsiDeploymentType -AddDetectionClause $AdditionalDetectionMethods -RemoveDetectionClause $DetectionLogicalName
            }

            Write-Host "Done adding deployment Type" -ForegroundColor Green
            $true
        }
        catch{
            Write-Host "Failed to adding deployment Type" -ForegroundColor Red
            $false    
            Write-Error $_
        }
        Set-Location $ReturnLocaltion.Path
    }
}
function New-RCMAppVDeploymentType {
    param(
        [Parameter(Mandatory=$true)]$APPName,
        [Parameter(Mandatory=$false)][string]$DeploymentTypeName="",
        [Parameter(Mandatory=$true)]$ContentLocation,
        $SlowNetworkDeploymentMode='Download',
        $FastNetworkDeploymentMode='Download',
        $ContentFallback=$true,
        $CloseExecutables
    )
        if(!$DeploymentTypeName){
            $DeploymentTypeName = $APPName
        }

        if($RCMisConnected -or $(Connect-RCM)){
        
        #this bit works best running from a file type location
        $ReturnLocaltion = Get-Location
        Set-Location $env:TEMP

        $Item0 = Get-Item -Path $ContentLocation
        if($Item0.PSIsContainer){
            $ContentDirectory = $Item0.FullName
            $AppVFile = @(Get-ChildItem -Path $ContentDirectory | ?{$_.Extension -eq ".appv"})[0].fullname
        }
        else{
            $AppVFile = $Item0.FullName
            $ContentDirectory = $Item0.Directory.FullName
        }
        Set-Location $ReturnLocaltion.Path
        

        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        $commands = @{
            ApplicationName=$APPName
            DeploymentTypeName=$DeploymentTypeName
            ContentLocation=$AppVFile 
            ContentFallback=$ContentFallback
            SlowNetworkDeploymentMode=$SlowNetworkDeploymentMode
            FastNetworkDeploymentMode=$FastNetworkDeploymentMode
            Force=$true
        }

        try{

            $deploymentType = Add-CMAppv5XDeploymentType @commands

            Write-Host "Done adding deployment Type" -ForegroundColor Green
            $true
        }
        catch{
            Write-Host "Failed to adding deployment Type" -ForegroundColor Red
            $false    
            Write-Error $_
        }

        Set-Location $ReturnLocaltion.Path
    }
}
Function New-RCMDeploymentTypeFromTemplate {
    param($Template,$Application)
    #Attempts to create a detection method
    if(($Template.DeploymentType.DetectionMethod) -and ($Template.DeploymentType.Tecnology -ne "App-V")){
        [array]$DetectionMethod = foreach($Detection in $Template.DeploymentType.DetectionMethod){
            try{
                if($Detection.Type -eq "App-V / will do later"){
                    New-RCMDetectionMethod -Type WindowsInstaller -Instructions @{"ProductCode"='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'}
                }
                else{
                    $DetectionSplat = $Detection | ConvertTo-Hashtable -RemoveNull
                    New-RCMDetectionMethod @DetectionSplat
                }

            }
            catch{
                Write-Host "Fail $($Detection.Type)"
            }
        }
    }

    
    #copys content to the content location
    Set-Location $env:TEMP
    if($Template.DeploymentType.ContentLocation -and $Template.DeploymentType.CopyFrom){
        if($Template.DeploymentType.CopyFrom -match "^\.\\"){
            $Template.DeploymentType.CopyFrom = $Template.DeploymentType.CopyFrom -replace "^\.",$PSScriptRoot
            Write-Host $Template.DeploymentType.CopyFrom -ForegroundColor Cyan
        }
        if($Template.DeploymentType.ContentLocation -ne $Template.DeploymentType.CopyFrom){
            if(!(Test-Path $Template.DeploymentType.ContentLocation)){
                $Folder = New-Item -Path $Template.DeploymentType.ContentLocation -ItemType directory -Force
            }
            elseif($Template.Name -ne $Application.LocalizedDisplayName){
                $Folder0 = Get-Item -Path $Template.DeploymentType.ContentLocation
                do {$NewFolderName = Get-IncrementedPackageVersion -Name $Folder0.Name}while(Test-Path "$($Folder0.Parent.FullName)\$NewFolderName")
                $Folder = New-Item -Path "$($Folder0.Parent.FullName)\$NewFolderName" -ItemType directory 
            }
            else{
                $Folder = Get-Item -Path $Template.DeploymentType.ContentLocation
            }
            $ComyFromObject = Get-Item -Path $Template.DeploymentType.CopyFrom
            if(Test-Path $Template.DeploymentType.CopyFrom){
                if($ComyFromObject.Extension -eq '.zip'){
                    Zip-Unzip -Unzip -ZipPath $Template.DeploymentType.CopyFrom -FolderPath $Folder.FullName
                }
                ROBOCOPY $Template.DeploymentType.CopyFrom $($Folder.FullName) /E /XO | Out-Null
            }
        }
    }
    Set-Location $RCMSiteCode
    #Prepare The Splat
    
    if($Template.DeploymentType.Tecnology -eq "App-V"){
        $Params = 'DeploymentTypeName','ContentLocation','SlowNetworkDeploymentMode','FastNetworkDeploymentMode','ContentFallback'
        $DeploymentTypeSplat = $Template.DeploymentType | Select-Object -Property $Params | ConvertTo-Hashtable -RemoveNull
    }
    else{
        $DeploymentTypeSplat = $Template.DeploymentType | Select-Object -Property *  -ExcludeProperty "DetectionMethod","Tecnology","CopyFrom" | ConvertTo-Hashtable -RemoveNull
    }

    if($DeploymentTypeSplat.CloseExecutables){
        $DeploymentTypeSplat.CloseExecutables = $DeploymentTypeSplat.CloseExecutables -split "," | %{$_.trim().trim('"').trim("'")}
    }
    
    $DtName = if($Template.DeploymentType.DeploymentTypeName){$Template.DeploymentType.DeploymentTypeName}else{$Application.LocalizedDisplayName}
    Write-Host $Application.LocalizedDisplayName -ForegroundColor Cyan
    $ExistingDeploymentType = Get-RCMDeploymentType -DeploymentTypeName $DtName -ApplicationName $Application.LocalizedDisplayName
    if($ExistingDeploymentType){
        Write-Host "$DtName Deployment type Already Exists" -ForegroundColor Yellow
        $NewName = $DtName

        do{
            $NewName = Get-IncrementedPackageVersion -Name $NewName
        }while(Get-RCMDeploymentType -DeploymentTypeName $NewName -ApplicationName $Application.LocalizedDisplayName)
        do{
            Write-Host "1: Attempt to Use Existing DeploymentType" -ForegroundColor Yellow
            Write-Host "2: Create as $NewName" -ForegroundColor Yellow
            Write-Host "3: Cancel Operation" -ForegroundColor Yellow
    
            $Read = Read-Host -Prompt "?:"
        }while(@("1","2","3") -notcontains $Read)
        if($Read -eq "1"){
            #$DtName = $NewName
            $DeploymentTypeObject = $ExistingDeploymentType
        }
        elseif($Read -eq "2"){
            $DtName = $NewName
            $DeploymentTypeSplat.DeploymentTypeName = $DtName
            if($Template.DeploymentType.Tecnology -eq "Script"){
                $DeploymentTypeObject = New-RCMScriptDeploymentType -APPName $Application.LocalizedDisplayName -DetectionMethods $DetectionMethod @DeploymentTypeSplat
            }
            elseif($Template.DeploymentType.Tecnology -eq "Msi"){
                $DeploymentTypeObject = New-RCMMSIDeploymentType -APPName $Application.LocalizedDisplayName -DetectionMethods $DetectionMethod @DeploymentTypeSplat
            }
            elseif($Template.DeploymentType.Tecnology -eq "App-V"){
                $DeploymentTypeObject = New-RCMAppVDeploymentType -APPName $Application.LocalizedDisplayName @DeploymentTypeSplat
            }
        }
        elseif($Read -eq "3"){
            return
        }
        else{Write-Error "How did this happen?"}
    }
    else{
        $DeploymentTypeSplat.DeploymentTypeName = $DtName
        if($Template.DeploymentType.Tecnology -eq "Script"){
            $DeploymentTypeObject = New-RCMScriptDeploymentType -APPName $Application.LocalizedDisplayName -DetectionMethods $DetectionMethod @DeploymentTypeSplat
        }
        elseif($Template.DeploymentType.Tecnology -eq "Msi"){
            #$DeploymentTypeObject = New-RCMMSIDeploymentType -APPName $Application.LocalizedDisplayName -DetectionMethods $DetectionMethod @DeploymentTypeSplat
            if($DetectionMethod){$DeploymentTypeObject = New-RCMMSIDeploymentType -APPName $Application.LocalizedDisplayName @DeploymentTypeSplat -DetectionMethods $DetectionMethod}
            else{$DeploymentTypeObject = New-RCMMSIDeploymentType -APPName $Application.LocalizedDisplayName @DeploymentTypeSplat}
        }
        elseif($Template.DeploymentType.Tecnology -eq "App-V"){
            #$DeploymentTypeSplat
            $DeploymentTypeObject = New-RCMAppVDeploymentType -APPName $Application.LocalizedDisplayName @DeploymentTypeSplat
        }
    }
    $DeploymentTypeObject 
    
}
function New-RCMScriptDeploymentType {
    param(
        [Parameter(Mandatory=$true)]$APPName,
        [Parameter(Mandatory=$false)]$DeploymentTypeName,
        [Parameter(Mandatory=$true)]$ContentLocation,
        [Parameter(Mandatory=$true)]$InstallScript,
        [Parameter(Mandatory=$false)]$RepairCommand,
        $ProductCode,
        [array]$DetectionMethods,
        $MaximumRuntimeMins=120,
        $EstimatedRuntimeMins=5,
        $UninstallScript,
        $UserInteractionMode="Hidden",
        $InstallationBehaviorType='InstallForSystem',
        $SlowNetworkDeploymentMode='Download',
        $LogonRequirementType='WhetherOrNotUserLoggedOn',
        $ContentFallback=$true,
        $CloseExecutables
    )
    if($RCMisConnected -or $(Connect-RCM)){

        if (!(Test-Path "FileSystem::$ContentLocation")){
            Write-Host "ContentLocation is invalid" -ForegroundColor Yellow
            break
        }

        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        if(!$DeploymentTypeName){$DeploymentTypeName = $APPName}
        try{
            $commands = @{
                ApplicationName=$APPName
                DeploymentTypeName=$DeploymentTypeName
                ContentLocation=$ContentLocation
                SlowNetworkDeploymentMode=$SlowNetworkDeploymentMode
                MaximumRuntimeMins=$MaximumRuntimeMins
                EstimatedRuntimeMins=$EstimatedRuntimeMins
                Force=$true
                InstallCommand=$InstallScript
                UserInteractionMode=$UserInteractionMode
                EnableBranchCache=$true
                ContentFallback=$ContentFallback
                InstallationBehaviorType=$InstallationBehaviorType
            }
            if($UninstallScript){$commands += @{UninstallCommand=$UninstallScript}}
            if($RepairCommand){$commands.RepairCommand = $RepairCommand}
            if($DetectionMethods){
                if($DetectionMethods){$commands  += @{AddDetectionClause=$DetectionMethods[0].method}}
                if($DetectionMethods.Count -gt 1){
                    [array]$AdditionalDetectionMethods = $DetectionMethods[1..$($DetectionMethods.Count -1)]
                }
            }
            else{$commands  += @{ProductCode="{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}"}}

            $deploymentType = Add-CMScriptDeploymentType @commands
            #this alters some stuff
            $AppObject = Get-CMApplication $AppName
            $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)  #grabs the XML in a eradible form
            
            $DTXML = $AppXML.DeploymentTypes | ?{"$($_.Scope)/$($_.Name)" -eq $deploymentType.ModelName}
            $DTXML.Installer[0].PostInstallBehavior = 'NoAction' #CHECK
            
            #AddsExecutablesToClose
            try{
                #AddsExecutablesToClose
                if($CloseExecutables){
                    foreach($C in $CloseExecutables){
                        $ProcessInfo = [Microsoft.ConfigurationManagement.ApplicationManagement.ProcessInformation]::new()
                        $ProcessInfo.Name = $C
                        $DTXML.Installer[0].InstallProcessDetection.ProcessList.Add($ProcessInfo)
                    }
                }
            }
            catch{
                Write-Host "Unable to add the Close Executables Process List"
            }

            $AppObject.SDMPackageXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($AppXML)
            $AppObject.Put()

            foreach($ADT in $AdditionalDetectionMethods){
                $deploymentType | Set-CMScriptDeploymentType -AddDetectionClause $ADT.method
            }

            Write-Host "Done adding deployment Type" -ForegroundColor Green
            $true
        
        }
        catch{
            Write-Host "Failed to adding deployment Type" -ForegroundColor Red
            $false    
            Write-Error $_
        }
        Set-Location $ReturnLocaltion.Path
    }
}
Function New-RCMDeploymentTypeTemplate {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$ApplicationName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,Position=0)]$DeploymentTypeName
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode
        $appXML = Get-RCMAppXml -Appname $ApplicationName
        if(!$DeploymentTypeName){
            $DeploymentTypeName = $appXML.DeploymentTypes | %{$_.title}
        }
        foreach($D in $DeploymentTypeName){
            try{
                $DeploymentTypeObject = Get-CMDeploymentType -DeploymentTypeName $D -ApplicationName $ApplicationName
                $DeploymentTypeXML = $appXML.DeploymentTypes | ?{$_.title -eq $D}
                #[Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($DeploymentTypeObject.SDMPackageXML).DeploymentTypes | ?{$_.Title -eq $DeploymentTypeObject.LocalizedDisplayName}
            }
            catch{
                continue
            }

            $Splat = [ordered]@{
                ExpressionOperator='IsEquals'
                PropertyType='ProductVersion'
                ExpectedValue=''
                ProductCode=''
                Value=$true
            }
        
            #Add-CMScriptDeploymentType -LogonRequirementType OnlyWhenUserLoggedOn

            $DetectionMethod = New-Object -TypeName psobject -Property $([ordered]@{Type='WindowsInstaller';Instructions=$Splat})

            $DTout0 = [ordered]@{
                DeploymentTypeName=''
                ContentLocation=''
                ContentFallback=$DeploymentTypeXML.Installer.Contents[0].FallbackToUnprotectedDP
                SlowNetworkDeploymentMode=$DeploymentTypeXML.Installer.Contents[0].OnSlowNetwork.ToString()
                InstallationBehaviorType=''
                UserInteractionMode=$DeploymentTypeXML.Installer[0].UserInteractionMode.ToString()
                LogonRequirementType=''
                Tecnology=''
                InstallScript=''
                UninstallScript=''
                MsiName=''
                MsTName=''
                MaximumRuntimeMins=$DeploymentTypeXML.Installer.MaxExecuteTime
                EstimatedRuntimeMins=$DeploymentTypeXML.Installer.ExecuteTime
                DetectionMethod=$DetectionMethod
            }
            $DTout = New-Object -TypeName psobject -Property $DTout0

            if($DeploymentTypeXML.Installer.Contents[0].FallbackToUnprotectedDP){$DTout.ContentFallback = $true}
            if($DeploymentTypeXML.Installer.RequiresLogOn){$DTout.LogonRequirementType = 'OnlyWhenUserLoggedOn'}
            else{$DTout.LogonRequirementType = 'WhetherOrNotUserLoggedOn'}

            switch ($DeploymentTypeXML.Installer.ExecutionContext.ToString()){
                "System" {$DTout.InstallationBehaviorType  = 'InstallForSystem'}
                "User" {$DTout.InstallationBehaviorType  = 'InstallForUser'}
                "Any" {$DTout.InstallationBehaviorType  = 'InstallForSystemIfResourceIsDeviceOtherwiseInstallForUser'}
                Default {$DTout.InstallationBehaviorType  = 'InstallForSystem'}
            }


            $DTout
        }
        Set-Location $ReturnLocaltion
    }
}
function New-RCMDistrubutionTemplate {
    Param($ApplicationName)
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        $APP = Get-CMApplication -Name $ApplicationName
        $APPDPS = $APP | %{Get-WmiObject -Namespace "root\SMS\site_$($RCMSiteCodeRaw)" -Class SMS_DPGroupDistributionStatusDetails -ComputerName $ProviderMachineName -Filter "PackageID = '$($_.PackageID)'" -ErrorAction SilentlyContinue}
        $APPDPSNames = $APPDPS | %{$_.InsString2}
        $AllGroups = Get-CMDistributionPointGroup

        $AccountedForDPs = @()
        $OUT = New-Object -TypeName psobject
        $Groups = Foreach ($G in $AllGroups){
            $GroupServers = Get-CMDistributionPoint -DistributionPointGroupName $G.name
            $DeployedToGroup = $true
            foreach ($S in $GroupServers){
                if($APPDPSNames -notcontains $S.NALPath){$DeployedToGroup = $false}
            }
            if($DeployedToGroup){
                $G.Name
                $AccountedForDPs += $GroupServers | %{$_.NALPath}
            }
        }
        $OUT | Add-Member -NotePropertyName "DistributionPointGroups" -NotePropertyValue $Groups
        $OtherDPs = $APPDPSNames |?{$AccountedForDPs -notcontains $_} | %{if($_ -match '\"Display=([^\"]*)\"'){$Matches[1]}} | %{$_.TrimEnd('\\')}
        $OUT | Add-Member -NotePropertyName "DistributionPoints" -NotePropertyValue $OtherDPs
        $OUT
        Set-Location $ReturnLocaltion.Path
    }
}
Function New-RCMTemplate {
    param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$ApplicationName,
        [switch]$AppBLOCK,
        [switch]$DeploymentTypeBLOCK,
        [switch]$DistrubutionBLOCK,
        [switch]$CollectionBLOCK,
        [switch]$DeploymentBLOCK#,
        #[switch]$Average
    )
    Begin{
        if($RCMisConnected -or $(Connect-RCM)){
            $ReturnLocaltion = Get-Location
            Set-Location $RCMSiteCode
        }
        else{break}
    }
    process{
        $AppObject0 = Get-CMApplication -Name $ApplicationName 
        foreach($AppObject in $AppObject0){
            $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($AppObject.SDMPackageXML)
            $DeploymentTypes = $AppXML.DeploymentTypes | %{$_.title}
            
            Write-Host $AppObject.LocalizedDisplayName -ForegroundColor Green

            $Out = New-Object -TypeName psobject
            $Out | Add-Member -NotePropertyName "Name" -NotePropertyValue ""
            $Out | Add-Member -NotePropertyName "AddApplication" -NotePropertyValue $false
            $Out | Add-Member -NotePropertyName "AddDeploymentType" -NotePropertyValue $false
            $Out | Add-Member -NotePropertyName "DistributeSource" -NotePropertyValue $false
            $Out | Add-Member -NotePropertyName "AddCollection" -NotePropertyValue $false
            $Out | Add-Member -NotePropertyName "AddDeployment" -NotePropertyValue $false
            $Out | Add-Member -NotePropertyName "AddAdGroup" -NotePropertyValue $false

            if(!$AppBLOCK){
                $AppTemplateObject = New-RCMAppTemplate -ApplicationObject $AppObject
                $Out | Add-Member -NotePropertyName "App" -NotePropertyValue $AppTemplateObject
                if($AppTemplateObject){
                    $Out.AddApplication = $true
                }
            }
            if(!$DeploymentTypeBLOCK){
                $DeploymentTypeObject = $DeploymentTypes | %{New-RCMDeploymentTypeTemplate -ApplicationName $AppObject.LocalizedDisplayName -DeploymentTypeName $_}
                $Out | Add-Member -NotePropertyName "DeploymentType" -NotePropertyValue $DeploymentTypeObject
                if($DeploymentTypeObject){
                    $Out.AddDeploymentType = $true
                }
            }
            if(!$DistrubutionBLOCK){
                $DistrubutionTemplateObject = New-RCMDistrubutionTemplate -ApplicationName $AppObject.LocalizedDisplayName
                $Out | Add-Member -NotePropertyName "Distrubution" -NotePropertyValue $DistrubutionTemplateObject
                if($DistrubutionTemplateObject){
                    $Out.DistributeSource = $true
                }
            }
            if(!$CollectionBLOCK){
                if(!$DeploymentBLOCK){
                    $CollectionTemplateObject = New-RCMCollectionTemplate -ApplicationName $AppObject.LocalizedDisplayName -Deployment
                    $Out | Add-Member -NotePropertyName "Collection" -NotePropertyValue $CollectionTemplateObject
                    if($CollectionTemplateObject){
                        $Out.AddCollection = $true
                        $Out.AddDeployment = $true
                    }
                    if($CollectionTemplateObject.GroupOU){
                        $Out.AddAdGroup = $true
                    }
                }
                else{
                    $CollectionTemplateObject = New-RCMCollectionTemplate -ApplicationName $AppObject.LocalizedDisplayName
                    $Out | Add-Member -NotePropertyName "Collection" -NotePropertyValue $CollectionTemplateObject
                    $Out.AddCollection = $true
                    if($CollectionTemplateObject){
                        $Out.AddCollection = $true
                        $Out.AddDeployment = $true
                    }
                    if($CollectionTemplateObject.GroupOU){
                        $Out.AddAdGroup = $true
                    }
                }
            }
            $Out
        }
    }
    end{
        Set-Location $ReturnLocaltion
    }
}
Function New-RcmTemplateFromSourceFile {
    param(
        $MainInstaller,
        $Template
    )
    
    $ReturnLocaltion = Get-Location
    Set-Location $env:TEMP
    $MainInstallerItem = Get-Item -Path $MainInstaller
    $Files = Get-ChildItem -Path $MainInstallerItem.Directory.FullName

    [array]$probableName = foreach ($F in $Files | Sort-Object -Property LastWriteTime -Descending){
        if($F.name -match "\d+\.\d+"){
            $F.name -replace "\.[^\.]+$",""
        }
    }
    if($probableName.Count -ge 1){
        $probableName = $probableName[0]
    }
    

    $Template.DeploymentType | Add-Member -NotePropertyName "CopyFrom" -NotePropertyValue $MainInstallerItem.Directory.FullName -Force
    $Template.DeploymentType | Add-Member -NotePropertyName "ContentLocation" -NotePropertyValue $MainInstallerItem.Directory.FullName -Force

    $Template.Name = $probableName
    #Write-Host "$($MainInstallerItem.Directory.FullName)   $($MainInstallerItem.FullName)" -ForegroundColor Yellow
    $InstallDetails = get-RcmPackageDetails -ContentPath $MainInstallerItem.Directory.FullName -Installer $MainInstaller

    $Template.DeploymentType.Tecnology = $InstallDetails.Technology
    
    $Template.DeploymentType | Add-Member -NotePropertyName "MsiName" -NotePropertyValue $InstallDetails.msi.Name -Force
    $Template.DeploymentType | Add-Member -NotePropertyName "MsTName" -NotePropertyValue $InstallDetails.mst.Name -Force
    $Template.DeploymentType | Add-Member -NotePropertyName "InstallScript" -NotePropertyValue $('"' + $InstallDetails.InstallScript.Name + '"') -Force
    $Template.DeploymentType | Add-Member -NotePropertyName "UninstallScript" -NotePropertyValue $('"' + $InstallDetails.UnInstallScript.name + '"') -Force

    #if($InstallDetails.Technology -eq "msi"){$Template.DeploymentType.MsiName = $InstallDetails.Installer.Name}
    #if($InstallDetails.Technology -eq "script"){$Template.DeploymentType.InstallScript = $InstallDetails.InstallScript.Name}
    #if($InstallDetails.Technology -eq "App-V"){$Template.DeploymentType.AppVName = $InstallDetails.Installer.Name}

    if($Template.DeploymentType.MsiName){
        $MSIProperties = Get-MSIProperty -Path "$($MainInstallerItem.Directory.FullName)\$($Template.DeploymentType.MsiName)"
        
        $DeploymentType = New-Object -TypeName psobject
        #$DeploymentType | Add-Member -NotePropertyName Type -NotePropertyValue "WindowsInstaller"
        $DeploymentType | Add-Member -NotePropertyName ProductCode -NotePropertyValue $MSIProperties.ProductCode
        $DeploymentType | Add-Member -NotePropertyName ExpectedValue -NotePropertyValue $MSIProperties.ProductVersion
        $DeploymentType | Add-Member -NotePropertyName Operator -NotePropertyValue "Equals"

        $Template.DeploymentType.DetectionMethod.Type = "WindowsInstaller"

        $Template.DeploymentType.DetectionMethod.Instructions = $DeploymentType
        #.Instructions.ExpectedValue = $MSIProperties.ProductVersion
    }
    if($InstallDetails.Technology -eq "App-V"){
        $Template.DeploymentType.DetectionMethod.Type = "App-V / will do later"
        $Template.DeploymentType.DetectionMethod.Instructions = $null
    }
   
    $Template

}
Function Add-RCMSupersedingDeploymentType  {
    param(
        [Parameter(Mandatory=$true)][string]$SupersedingApplicationName,
        [Parameter(Mandatory=$false)][string]$SupersedingDeploymentTypeName,
        [Parameter(Mandatory=$true)][string]$SupersededApplicationName,
        [switch]$UninstallSuperseded
    )
    Write-Host "Superseding $SupersedingApplicationName / $SupersedingDeploymentTypeName / $SupersededApplicationName / $UninstallSuperseded" -ForegroundColor Cyan
    if($RCMisConnected -or $(Connect-RCM)){
        if($SupersedingDeploymentTypeName){
            $SupersedingDeploymentType = Get-CMDeploymentType -ApplicationName $SupersedingApplicationName -DeploymentTypeName $SupersedingDeploymentTypeName
        }
        else{
            $SupersedingDeploymentType = @(Get-CMDeploymentType -ApplicationName $SupersedingApplicationName)[0]
        }
        $SupersededDeploymentTypeS = Get-CMDeploymentType -ApplicationName $SupersededApplicationName

        if($SupersedingDeploymentType -and $SupersededDeploymentTypeS){
            foreach($SSDT in $SupersededDeploymentTypeS){
                $SPLAT = @{
                    SupersedingDeploymentType=$SupersedingDeploymentType
                    SupersededDeploymentType=$SSDT
                }
                if($UninstallSuperseded){$SPLAT.IsUninstall = $true}
                Add-CMDeploymentTypeSupersedence @SPLAT 
            }
        }
    }
}
Function Add-RCMDeploymentTypeDependency {
    param(
        [Parameter(Mandatory=$true)][string]$ChildApplicationName,
        [Parameter(Mandatory=$true)][string]$DependencyApplicationName,
        [Parameter(Mandatory=$false)][string]$DeploymentTypeDependencyGroupName,
        [switch]$InstallDependency
    )
    Write-Host "Dependency $ChildApplicationName / $DependencyApplicationName / $DeploymentTypeDependencyGroupName / $InstallDependency" -ForegroundColor Cyan
    if($RCMisConnected -or $(Connect-RCM)){
        $ChildDeploymentTypes = Get-CMDeploymentType -ApplicationName $ChildApplicationName
        #Get-CMDeploymentType -ApplicationName
        $DependencyDeploymentTypes = Get-CMDeploymentType -ApplicationName $DependencyApplicationName

        $SPLAT = @{
            DeploymentTypeDependency=$DependencyDeploymentTypes
        }
        if($InstallDependency){$SPLAT.IsAutoInstall = $true}

        if(!$DeploymentTypeDependencyGroupName){$DeploymentTypeDependencyGroupName = $DependencyApplicationName}
        
        $ChildDeploymentTypes | New-CMDeploymentTypeDependencyGroup -GroupName $DeploymentTypeDependencyGroupName | Add-CMDeploymentTypeDependency @SPLAT
    }
}
Function New-RCMSuperDependFromTemplate {
    param($Template,$Application,$DeploymentType)
    #Supersedance
    foreach($DT in $DeploymentType){
        foreach($S in $Template.SupersedanceDependency.Supersedance){
            if($S.Supersede -match "\S"){
                $SPLAT = @{
                    SupersedingApplicationName=$Application.LocalizedDisplayName
                    SupersedingDeploymentTypeName=$DT.LocalizedDisplayName
                    SupersededApplicationName=$S.Supersede
                }
                if($S.Uninstall){
                    $SPLAT.UninstallSuperseded=$true
                }
                Add-RCMSupersedingDeploymentType @SPLAT
            }
        }
    }
    foreach($Dep in $Template.SupersedanceDependency.Dependency){
        if($Dep.Dependency){
            $SPLAT = @{
                ChildApplicationName=$Application.LocalizedDisplayName
                DependencyApplicationName=$Dep.Dependency
            }
            if($Dep.AutoInstall){
                $SPLAT.InstallDependency = $true
            }
            Add-RCMDeploymentTypeDependency @SPLAT
        }
    }
}
#endregion

#region GUI Functions
Function Invoke-RWPFTreeBrowse {
param($Root,$Object=$(Get-ChildItem $Root -Directory -Force -Recurse))
[xml]$TreeXml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Rory's SCCM Importer" WindowStartupLocation = "CenterScreen"
    SizeToContent = "WidthAndHeight" 
    ShowInTaskbar = "True" 
    Background = "White" 
    ResizeMode = "NoResize"
>
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="MainStackPanel">
        <TreeView Name="MotherTree">
        </TreeView>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "SelectItem" Height = "26" Content = '  Select Item ' />
            <Button x:Name = "Cancel" Height = "26" Content = ' Cancel ' />
        </StackPanel>
    </StackPanel>
</Window>
"@

    Add-Type -AssemblyName PresentationFramework
    $TreeReader=(New-Object System.Xml.XmlNodeReader $TreeXml)
    $TreeWindow=[Windows.Markup.XamlReader]::Load($TreeReader)

    $Root = $Root.TrimEnd("\")

    function Add-RWFPTreeChild {
        param(
            $mother,
            [Parameter(Mandatory=$true,ValueFromPipeline=$true)]$Child
        )
        begin{}
        process{
            $ChildItem = New-Object -TypeName System.Windows.Controls.TreeViewItem
            $ChildItem.Header = $Child.Name
            $ChildItem.Tag = $Child
            $mother.AddChild($ChildItem)
            $ChildItem
        }
        end{}
    }
    function add-RWFPBuildTree {
        param($MotherTree,$TreeInput,$Root)
        $Children = $TreeInput | ?{$_.Parent.FullName -eq $Root}  | Sort-Object -Property name
        if(!$Children){
            $ToMatch = [regex]::Escape($Root)
            $ImpliedChildren = $TreeInput | %{if($_.Parent.FullName -match "^$ToMatch\\[^\\]+"){$Matches[0]}}  | Sort-Object -Unique
            if($ImpliedChildren){
                $Children = $ImpliedChildren | %{New-Object psobject -Property @{fullname=$_;Name=$($_.split("\")[-1])}}
            }
        }
        foreach($C in $Children){
            $ChildTree = Add-RWFPTreeChild -Child $C -mother $MotherTree
            #Write-Host $C.FullName -ForegroundColor Yellow
            add-RWFPBuildTree -MotherTree $ChildTree -TreeInput $($TreeInput | ?{$_.fullname -like "$($C.FullName)\*"}) -Root $C.FullName
        }
    }

    $MainStackPanel  = $TreeWindow.FindName('MainStackPanel')
    $MotherTree  = $TreeWindow.FindName('MotherTree')
    $SelectItem  = $TreeWindow.FindName('SelectItem')
    $Cancel  = $TreeWindow.FindName('Cancel')

    #$RootItem = $Object
    $TreeInput = $Object

    $Cancel.add_click({
        $TreeWindow.Close()
    })
    $SelectItem.add_click({

        if($MotherTree.SelectedItem.tag.DistinguishedName){
            $Global:Output_585f7b4ae58348a1b0c7faec37fed853 = $MotherTree.SelectedItem.tag.DistinguishedName
        }
        else{
            $Global:Output_585f7b4ae58348a1b0c7faec37fed853 = $MotherTree.SelectedItem.tag.fullname
        }
        Write-Host $MotherTree.SelectedItem.tag.fullname -ForegroundColor Cyan
        $TreeWindow.Close()
    })

    add-RWFPBuildTree -MotherTree $MotherTree -TreeInput $TreeInput -Root $Root

    $async = $TreeWindow.Dispatcher.InvokeAsync({$TreeWindow.ShowDialog()})
    $async.Wait() | Out-Null

    $Global:Output_585f7b4ae58348a1b0c7faec37fed853
}
function New-TextBoxPanel {
    param(
        [string]$LabelName,
        [string]$LogicalName,
        [string]$Content,
        [string]$Key,
        $GRandMother
    )
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    $SPMother.Name = "$LogicalName$('SP')"

    $LB = New-Object System.Windows.Controls.Label
    $LB.HorizontalContentAlignment = "Right"
    $LB.Content = "$LabelName"
    $LB.Name = "$LogicalName$("LB")"
    $LB.Width = 120
    $SPMother.AddChild($LB)

    $TB = New-Object System.Windows.Controls.Textbox
    $TB.Width = 600
    $TB.Name = "$LogicalName$("TB")"
    $TB.Tag = $Key
    $TB.Text = $Content

    $SPMother.AddChild($TB)
    $GRandMother.AddChild($SPMother)
    
    $HashName = $this.parent.parent.tag
    if($HashName){
        $HashObject = Get-Variable -Name $HashName -ValueOnly
        $TB.Text = $HashObject.$LogicalName.Content
    }
}
function New-TextboxBrowsePanel {
    param(
        [string]$LabelName,
        [string]$LogicalName,
        [string]$Content,
        [string]$Key,
        $ObjectBlock,
        $Root,
        $GRandMother
    )
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    $SPMother.Name = "$LogicalName$('SP')"

    $LB = New-Object System.Windows.Controls.Label
    $LB.HorizontalContentAlignment = "Right"
    $LB.Content = "$LabelName"
    $LB.Name = "$LogicalName$("LB")"
    $LB.Width = 120
    $SPMother.AddChild($LB)

    $TB = New-Object System.Windows.Controls.Textbox
    $TB.Width = 500
    $TB.Name = "$LogicalName$("TB")"
    $TB.Tag = $Key
    $TB.Text = $Content
    $SPMother.AddChild($TB)

    $BT = New-Object System.Windows.Controls.Button
    $BT.Width = 100
    $BT.Name = "$LogicalName$("BT")"
    $BT.Tag = @{Root=$Root;Block=$ObjectBlock}
    $BT.Content = "Browse"
    $BT.Add_Click({
        #Invoke browse dialog
        #$Object = $(Get-RCMChildFolder -Root -FolderType Application -Recurse)
        
        Write-Host $this.tag[1]
        #$this.parent.children[1].Text = Invoke-RWPFTreeBrowse -Object $(Get-RCMChildFolder -Root -FolderType Application -Recurse) -Root $this.tag[0]
        #$Output = Invoke-RWPFTreeBrowse -Root $this.tag[0] -Object $(Get-ChildItem -Path C:\Scripts -Recurse) #$(&$this.tag[1])
        try{
            
            $Object = @(&$this.tag.Block)
           
            $Root = if($this.tag.Root){$this.tag.Root}else{$($Object[0].FullName -split "\\")[0]}
            Write-Host "InvokeRoot: $Root" -ForegroundColor Yellow
            $Output = Invoke-RWPFTreeBrowse -Root $Root -Object $Object
            Write-Host $Output -ForegroundColor Yellow
        }
        catch{
            Write-Host "Can't generate tree" -ForegroundColor Yellow
        }
        $this.parent.children[1].Text = $Output
        #$this.parent.children[1].Text = $Content
    })
    $SPMother.AddChild($BT)
    
    $GRandMother.AddChild($SPMother)
    
    $HashName = $this.parent.parent.tag
    if($HashName){
        $HashObject = Get-Variable -Name $HashName -ValueOnly
        $TB.Text = $HashObject.$LogicalName.Content
    }
}
function New-ComboBoxPanel {
    param(
        [string]$LabelName,
        [string]$LogicalName,
        [array]$CBItems,
        [string]$Key,
        $Index=0,
        $GRandMother,
        [switch]$Passthrough
    )

    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    $SPMother.Name = "$LogicalName$('SP')"

    $LB= New-Object System.Windows.Controls.Label
    $LB.HorizontalContentAlignment = "Right"
    $LB.Content = "$LabelName"
    $LB.Name = "$LogicalName$("LB")"
    $LB.Width = 120
    $SPMother.AddChild($LB) | Out-Null
    $CB = New-Object System.Windows.Controls.ComboBox
    $CB.Width = 600
    $CB.Name = "$LogicalName$("CB")"
    $CB.Tag = $Key
    $CB.StaysOpenOnEdit = $true
    foreach($C in $CBItems){
        $N = New-Object System.Windows.Controls.ComboBoxItem
        $N.Content = $C
        $CB.Items.Add($N) | Out-Null
    }

    $SPMother.AddChild($CB) | Out-Null
    $GRandMother.AddChild($SPMother) | Out-Null

    $CB.SelectedIndex = $Index
    if($Passthrough){$CB}
}
function Add-SP {
    param($Mother,$Hash,$Exclude,[switch]$PassThru)
    foreach($K in $($Hash.keys |?{$Exclude -notcontains $_})){
        if($Hash.$K.Type -eq "Textbox"){
            $Child = New-TextBoxPanel -LabelName $Hash.$K.LabelName -Key $K -LogicalName $Hash.$K.LogicalName -Content $Hash.$K.Content -GRandMother $Mother
        }
        elseif($Hash.$K.Type -eq "ComboBox"){
            $Child = New-ComboBoxPanel -LabelName $Hash.$K.LabelName -Key $K -LogicalName $Hash.$K.LogicalName -CBItems $Hash.$K.CBItems -Index $Hash.$K.Index -GRandMother $Mother
        }
        if($Hash.$K.Type -eq "TextboxBrowse"){
            $Child = New-TextboxBrowsePanel -LabelName $Hash.$K.LabelName -Key $K -LogicalName $Hash.$K.LogicalName -Content $Hash.$K.Content -GRandMother $Mother -Root $Hash.$K.Root -ObjectBlock $Hash.$K.ObjectBlock #-BrowseObject $Hash.$K.BrowseObject
        }
        if($PassThru){$Child}
    }
}
function New-DeploymentItem {
    param($Mother,$Type,$Checkbox)
    
    $SPMother = New-Object System.Windows.Controls.StackPanel
    $SPMother.Orientation = 'Horizontal'
    $SPMother.Tag = $type
    
    $LB = New-Object System.Windows.Controls.Label
    $LB.HorizontalContentAlignment = "Right"
    $LB.Content = " "
    $LB.Width = 150
    $SPMother.AddChild($LB)

    
    #<ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="DistributionPointGroupCB" Height = "26" Width = "570" KeyboardNavigation.TabIndex="0"/>
    $CBs = New-Object System.Windows.Controls.ComboBox
    
    $CBs.IsTextSearchEnabled = $false
    $CBs.IsEditable = $true
    $CBs.TabIndex = 0
    $CBs.Width = if($Checkbox){470}else{540}
    Set-RComboBoxOptions -ComboBox $CBs -ItemList $Mother.tag -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}
    #$CBS.add_DropDownOpened({
    #    DropdownRefineAction
    #})
    $SPMother.AddChild($CBs)


    

    if($Checkbox){
        $CheckboxBX = New-Object System.Windows.Controls.CheckBox
        $CheckboxBX.Content = $Checkbox
        $CheckboxBX.IsChecked = $true
        $CheckboxBX.Width = 70
        $SPMother.AddChild($CheckboxBX)
    }

    $RB = New-Object System.Windows.Controls.Button
    $RB.Width = 30
    $RB.Content = "-"
    $SPMother.AddChild($RB)

    $RB.Add_Click({
        $this.Parent.Parent.Children.Remove($this.Parent)
    })
    $Button = $Mother.Children[-1]
    $Mother.Children.Remove($Button)
    $Mother.AddChild($SPMother)
    $Mother.AddChild($Button)
    #@($CBs.parent.parent.children[0].children[1].tag) | %{New-Object System.Windows.Controls.ComboBoxItem -Property @{Content=$_}} | %{$CBs.Items.Add($_)} | %{Out-Null}
    #$CBs.ItemsSource = @($CBs.parent.parent.children[0].children[1].tag)
}
function Update-PannelObject {
    param($Mother,[string]$PanelName,$Value)
    $Pannel = $Mother.Children | ?{$_.children[0].content -eq $PanelName}
    $Child = $Pannel.Children[1]
    if($Child.GetType().Name -eq "ComboBox"){
        #Write-Host "$Value" -ForegroundColor Yellow
        $Child.SelectedItem = $Child.Items | ?{$_.content -eq $Value}
    }
    elseif($Child.GetType().Name -eq "TextBox"){
        $Child.Text = $Value
    }
}
function Load-Template {
    param($Template)
    
    $ApplicationCH.IsChecked = $Template.AddApplication
    $DeploymentTypeCH.IsChecked = $Template.AddDeploymentType
    $DistributionCH.IsChecked = $Template.DistributeSource
    $CollectionCH.IsChecked = $Template.AddCollection
    $DeploymentCH.IsChecked = $Template.AddDeployment
    $ADGroupsCH.IsChecked = $Template.AddAdGroup

    #$MainOut | Add-Member -NotePropertyName "AddApplication" -NotePropertyValue $ApplicationCH.IsChecked
    #$MainOut | Add-Member -NotePropertyName "AddDeploymentType" -NotePropertyValue $DeploymentTypeCH.IsChecked
    #$MainOut | Add-Member -NotePropertyName "DistributeSource" -NotePropertyValue $DistributionCH.IsChecked
    #$MainOut | Add-Member -NotePropertyName "AddCollection" -NotePropertyValue $CollectionCH.IsChecked
    #$MainOut | Add-Member -NotePropertyName "AddDeployment" -NotePropertyValue $DeploymentCH.IsChecked
    #$MainOut | Add-Member -NotePropertyName "AddAdGroup" -NotePropertyValue $ADGroupsCH.IsChecked
    #App Section
    Update-PannelObject -Mother $ApplicationSP -PanelName "Application Name:" -Value $Template.Name
    $ApplicationHash.Keys | %{$ApplicationHash.$_} | ?{$_.LabelName -ne "Application Name:"} | %{if($Template.App.$($_.LogicalName)){Update-PannelObject -Mother $ApplicationSP -PanelName $_.LabelName -Value $Template.App.$($_.LogicalName)}}
    #DeploymetType Section
    $DeploymetTypeHash.Keys | %{$DeploymetTypeHash.$_} | %{if($Template.DeploymentType.$($_.LogicalName)){Update-PannelObject -Mother $DeploymentTypeDynamicSP -PanelName $_.LabelName -Value $Template.DeploymentType.$($_.LogicalName)}}
    if($Template.DeploymentType.Tecnology){$TecnologyCB.SelectedItem = $TecnologyCB.Items | ?{$_.content -eq $Template.DeploymentType.Tecnology}}
    else{$TecnologyCB.SelectedIndex = 0}

    $DC = 0
    foreach($TemplateDetection in $Template.DeploymentType.DetectionMethod){
        if(!($MasterDMSP.Children[$DC])){Add-DetectionMethod}
        
        $Connector = $MasterDMSP.Children[$DC].children | ?{$_.tag -eq "DetectionMethodConnector"}
        $DCBSP =  $MasterDMSP.Children[$DC].Children | ?{$_.tag -eq "TYPE"}
        $WISP =  $MasterDMSP.Children[$DC].Children | ?{$_.tag -eq "MSI"}
        $FSSP =  $MasterDMSP.Children[$DC].Children | ?{$_.tag -eq "FILE"}
        $REGSP =  $MasterDMSP.Children[$DC].Children | ?{$_.tag -eq "REG"}
        
        switch($TemplateDetection.Type){
            "WindowsInstaller"      {$DetectionMethodWI.Keys | %{$DetectionMethodWI.$_} | %{if($TemplateDetection.Instructions.$($_.LogicalName)){Update-PannelObject -Mother $WISP  -PanelName $_.LabelName -Value $TemplateDetection.Instructions.$($_.LogicalName)}}}
            "FilesyStem"            {$DetectionMethodFS.Keys | %{$DetectionMethodFS.$_} | %{if($TemplateDetection.Instructions.$($_.LogicalName)){Update-PannelObject -Mother $FSSP  -PanelName $_.LabelName -Value $TemplateDetection.Instructions.$($_.LogicalName)}}}
            "Registry"              {$DetectionMethodREG.Keys| %{$DetectionMethodREG.$_}| %{if($TemplateDetection.Instructions.$($_.LogicalName)){Update-PannelObject -Mother $REGSP -PanelName $_.LabelName -Value $TemplateDetection.Instructions.$($_.LogicalName)}}}
            #"App-V / will do later" {$DetectionMethodREG.Keys| %{$DetectionMethodREG.$_}| %{if($TemplateDetection.Instructions.$($_.LogicalName)){Update-PannelObject -Mother $REGSP -PanelName $_.LabelName -Value $TemplateDetection.Instructions.$($_.LogicalName)}}}
        }
        
        if($TemplateDetection.Type){
            $DCBSP.Children[1].SelectedItem = $DCBSP.Children[1].Items | ?{$_.content -eq $TemplateDetection.Type}
            Write-Verbose $DCBSP.Children[1].SelectedItem.content
            
        }
        else{$DetectionTypeCB.SelectedIndex = 0}
        if($TemplateDetection.Connector -and ($DC -gt 0)){$Connector.Children[1].SelectedItem = $Connector.Children[1].Items | ?{$_.content -eq $TemplateDetection.Connector}}
        
        $DC++
    }

    #Supersedance
    while (($($SupersedenceSP.Children.Count -1) -le @($Template.SupersedanceDependency.Supersedance).Count) -or  ($SupersedenceSP.Children.Count -lt 2)){New-DeploymentItem -Mother $SupersedenceSP -Type "Supersedance" -Checkbox "Uninstall"}
    while (($($SupersedenceSP.Children.Count -1) -gt @($Template.SupersedanceDependency.Supersedance).Count) -and ($SupersedenceSP.Children.Count -gt 2)){$SupersedenceSP.Children.RemoveAt($SupersedenceSP.Children.count -2)}
    for ($N = 0; $N -lt $Template.SupersedanceDependency.Supersedance.Count; $N++){
        $SupersedenceSP.Children[$N].Children[1].Text = @($Template.SupersedanceDependency.Supersedance[$N].Supersede)
        $SupersedenceSP.Children[$N].Children[2].ischecked = @($Template.SupersedanceDependency.Supersedance[$N].Uninstall)
    }
    #Dependency
    while (($($DependencySP.Children.Count -1) -le @($Template.SupersedanceDependency.Dependency).Count) -or  ($DependencySP.Children.Count -lt 2)){New-DeploymentItem -Mother $DependencySP -Type "Dependency" -Checkbox "Install"}
    while (($($DependencySP.Children.Count -1) -gt @($Template.SupersedanceDependency.Dependency).Count) -and ($DependencySP.Children.Count -gt 2)){$DependencySP.Children.RemoveAt($DependencySP.Children.count -2)}
    for ($N = 0; $N -lt $Template.SupersedanceDependency.Dependency.Count; $N++){
        $DependencySP.Children[$N].Children[1].Text = @($Template.SupersedanceDependency.Dependency[$N].Dependency)
        $DependencySP.Children[$N].Children[2].ischecked = @($Template.SupersedanceDependency.Dependency[$N].AutoInstall)
    }

    #Distribution
    #sets the number of groups equal to the template
    while (($($DPGSP.Children.Count -1) -le @($Template.Distrubution.DistributionPointGroups).Count) -or  ($DPGSP.Children.Count -lt 2)){New-DeploymentItem -Mother $DPGSP -Type "DPG"}
    while (($($DPGSP.Children.Count -1) -gt @($Template.Distrubution.DistributionPointGroups).Count) -and ($DPGSP.Children.Count -gt 2)){$DPGSP.Children.RemoveAt($DPGSP.Children.count -2)}
    for ($N = 0; $N -lt $Template.Distrubution.DistributionPointGroups.Count; $N++){$DPGSP.Children[$N].Children[1].Text = @($Template.Distrubution.DistributionPointGroups)[$N]}
    
    #$DPGSP.Children[6].Children[1].GetType()
    while ($(($DPSP.Children.Count -1) -le @($Template.Distrubution.DistributionPoints).Count) -or  ($DPSP.Children.Count -lt 2)){New-DeploymentItem -Mother $DPSP -Type "DP"}
    while ($(($DPSP.Children.Count -1) -gt @($Template.Distrubution.DistributionPoints).Count) -and ($DPSP.Children.Count -gt 2)){$DPSP.Children.RemoveAt($DPSP.Children.Count -2)}
    for ($N = 0; $N -lt $Template.Distrubution.DistributionPoints.count; $N++){$DPSP.Children[$N].Children[1].Text = @($Template.Distrubution.DistributionPoints)[$N]}

    #Collection
    while (($($CollectionMasterSP.Children.Count) -le @($Template.Collection).Count) -or  ($CollectionMasterSP.Children.Count -lt 1)){New-CollectionDeploymentSP -HashTemplateC $CollectionHASHTemplate -HashTemplateD $DeploymentHASHTemplate -Mother $CollectionMasterSP | Out-Null}
    while (($($CollectionMasterSP.Children.Count) -gt @($Template.Collection).Count) -and ($CollectionMasterSP.Children.Count -gt 1)){$CollectionMasterSP.Children.RemoveAt(0)}

    for ($N = 0; $N -lt @($Template.Collection).Count; $N++){
        #$Template = $Template.Collection[0]
        $CollectionHASHTemplate.Keys | %{$CollectionHASHTemplate.$_}| %{if(@($Template.Collection)[$N].$($_.LogicalName)){Update-PannelObject -Mother $CollectionMasterSP.Children[$N].Children[1] -PanelName $_.LabelName -Value @($Template.Collection)[$N].$($_.LogicalName)}}
        $DeploymentHASHTemplate.Keys | %{$DeploymentHASHTemplate.$_}| %{if(@($Template.Collection.Deployment)[$N].$($_.LogicalName)){Update-PannelObject -Mother $CollectionMasterSP.Children[$N].Children[3] -PanelName $_.LabelName -Value @($Template.Collection.Deployment)[$N].$($_.LogicalName)}}
    }

}
Function Import-JsonTemplate {
    param($Path,[switch]$Clear)
    $returnLocation = Get-Location
    Set-Location $env:TEMP
    $Template = Get-Content -Path $Path -raw | ConvertFrom-Json
    #If we are in a loading template situation we need to clear the old template
    Set-Location $returnLocation
    if($Clear){Start-Up -Clear}
    Load-Template -Template $Template
}
Function New-CollectionDeploymentSP {
    Param($HashTemplateC,$HashTemplateD,$Mother)
    
    $SPCollectionDeployment = New-Object System.Windows.Controls.StackPanel
    $SPCollectionDeployment.Orientation = 'Vertical'
    $SPCollection = New-Object System.Windows.Controls.StackPanel
    $SPCollection.Orientation = 'Vertical'
    $SPDeployment = New-Object System.Windows.Controls.StackPanel
    $SPDeployment.Orientation = 'Vertical'
    #$GCAC = $Global:CollectionArray.Count

    $RB = New-Object System.Windows.Controls.Button
    $RB.Width = 200
    $RB.Content = "Remove Collection (below)"
    $RB.Add_Click({$this.Parent.Parent.Children.Remove($this.Parent)})
    $SPCollectionDeployment.AddChild($RB)

    $CollectionHashName = [guid]::NewGuid().guid
    $DeploymentHashName = [guid]::NewGuid().guid

    $HASHC = $HashTemplateC
    $HASHD = $HashTemplateD

    #$Global:CollectionArray += @{Collection=$HASHC;Deployment=$HASHD}
    $DeploymentLabel = New-Object System.Windows.Controls.Label
    $DeploymentLabel.Content = "Deployment"    
    
    $SPCollectionDeployment.AddChild($SPCollection)
    $SPCollectionDeployment.AddChild($DeploymentLabel)
    $SPCollectionDeployment.AddChild($SPDeployment)
    
    Add-SP -Mother $SPCollection -Hash $HASHC

    Add-SP -Mother $SPDeployment -Hash $HASHD

    $Mother.AddChild($SPCollectionDeployment)
    $SPCollectionDeployment
}
Function StringConvert {
    param([Parameter(Mandatory=$true,ValueFromPipeline=$true)]$Object)
    #$Object = $MainOut 
    foreach ($P in $object.psobject.Properties.name){
        if($object.$P){
            if($object.$P.gettype().name -eq "string"){
                if($object.$P -match '^\d*$'){
                    $object.$P = $object.$P.ToInt32($null)
                }
                elseif($object.$P -eq "true"){
                    $object.$P = $true
                }
                elseif($object.$P -eq "false"){
                    $object.$P = $false
                }
            }
            
            elseif(($object.$P.gettype().name -eq "PSCustomObject") -or ($object.$P.gettype().name -eq 'Object[]')){
                $object.$P = foreach($OP in $object.$P){
                    #Write-Host "Go deeper" -ForegroundColor Cyan
                    StringConvert -Object $OP
                }
            }
        }
    }
    $Object
}
Function Export-Template {
    param([Parameter(Mandatory=$false)][string]$path,[switch]$Pass)
    #Template Objext equilivant

    $MainOut = New-Object -TypeName psobject

    #application
    $Application = New-Object -TypeName psobject
    foreach($C in $ApplicationSP.Children){
        if($C.Children[1].tag -eq "ApplicationName"){
            $MainOut | Add-Member -NotePropertyName "Name" -NotePropertyValue $C.Children[1].Text -Force
        }
        else{
            $Application | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].Text -Force
            #Write-Host $C.Children[1].Text -ForegroundColor Yellow
        }
    }

    $MainOut | Add-Member -NotePropertyName "AddApplication" -NotePropertyValue $ApplicationCH.IsChecked
    $MainOut | Add-Member -NotePropertyName "AddDeploymentType" -NotePropertyValue $DeploymentTypeCH.IsChecked
    $MainOut | Add-Member -NotePropertyName "DistributeSource" -NotePropertyValue $DistributionCH.IsChecked
    $MainOut | Add-Member -NotePropertyName "AddCollection" -NotePropertyValue $CollectionCH.IsChecked
    $MainOut | Add-Member -NotePropertyName "AddDeployment" -NotePropertyValue $DeploymentCH.IsChecked
    $MainOut | Add-Member -NotePropertyName "AddAdGroup" -NotePropertyValue $ADGroupsCH.IsChecked

    $MainOut | Add-Member -NotePropertyName "App" -NotePropertyValue $Application -Force

    $DeploymentType = New-Object -TypeName psobject
    $DeploymentType | Add-Member -NotePropertyName "Tecnology" -NotePropertyValue $TecnologyCB.SelectedItem.Content
    
    foreach($C in $DeploymentTypeDynamicSP.Children){if($C.Visibility.value__ -eq 0){$DeploymentType | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}}
    
    $DetectionMethods = foreach ($Detection in $MasterDMSP.Children){
        $Connector = $Detection.children | ?{$_.tag -eq "DetectionMethodConnector"}
        #$NotConnector = $Detection.children | ?{$_.tag -ne "DetectionMethodConnector"}
        
        $DetectionComboBox =  $Detection.children | ?{$_.tag -eq "TYPE"}

        $WISP =  $Detection.children | ?{$_.tag -eq "MSI"}
        $FSSP =  $Detection.children | ?{$_.tag -eq "FILE"}
        $REGSP =  $Detection.children | ?{$_.tag -eq "REG"}

        $DetectionOut0 = New-Object -TypeName psobject
        $DetectionOut0 | Add-Member -NotePropertyName "Type" -NotePropertyValue $DetectionComboBox.Children[1].Text
        if($Connector){$DetectionOut0 | Add-Member -NotePropertyName "Connector" -NotePropertyValue $Connector.children[1].Text}
        
        $DetectionOut1 = New-Object -TypeName psobject

        switch ($DetectionComboBox.Children[1].Text) {
            "WindowsInstaller" {foreach($C in $WISP.Children ){$DetectionOut1 | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}}
            "FileSystem"       {foreach($C in $FSSP.Children ){$DetectionOut1 | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}}
            "Registry"         {foreach($C in $REGSP.Children){$DetectionOut1 | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}}
        }
        $DetectionOut0 | Add-Member -NotePropertyName "Instructions" -NotePropertyValue $DetectionOut1 -Force
        $DetectionOut0
    }

    $DeploymentType | Add-Member -NotePropertyName "DetectionMethod" -NotePropertyValue $DetectionMethods -Force
    $MainOut | Add-Member -NotePropertyName "DeploymentType" -NotePropertyValue $DeploymentType -Force
    
    #Write-Host $SupersedenceSP.Children[0].Children[1].Text -ForegroundColor Cyan

    [array]$Supersedance = foreach($C in $SupersedenceSP.Children){
        if($C.Children[1].text){
            $OUT = New-Object -TypeName psobject
            $OUT | Add-Member -NotePropertyName "Supersede" -NotePropertyValue $C.Children[1].text
            #$OUT | Add-Member -NotePropertyName "SupersedingDT" -NotePropertyValue $C.Children[1].text
            $OUT | Add-Member -NotePropertyName "Uninstall" -NotePropertyValue $C.Children[2].IsChecked
            $OUT
        }
    }
    #$Supersedance | %{Write-Host $_ -ForegroundColor Green}
    [array]$Dependency = foreach($C in $DependencySP.Children){
        if($C.Children[1].text){
            $OUT = New-Object -TypeName psobject
            $OUT | Add-Member -NotePropertyName "Dependency" -NotePropertyValue $C.Children[1].text
            #$OUT | Add-Member -NotePropertyName "SupersedingDT" -NotePropertyValue $C.Children[1].text
            $OUT | Add-Member -NotePropertyName "AutoInstall" -NotePropertyValue $C.Children[2].IsChecked
            $OUT
        }
    }
    $SupersedanceDependency = New-Object -TypeName psobject
    $SupersedanceDependency | Add-Member -NotePropertyName 'Supersedance' -NotePropertyValue $Supersedance
    $SupersedanceDependency | Add-Member -NotePropertyName 'Dependency' -NotePropertyValue $Dependency
    #$Dependency | %{Write-Host $_ -ForegroundColor Yellow}

    $DistributionGroups = foreach($C in $DPGSP.Children){if($C.Children[1].text){$C.Children[1].text}}
    $DistributionPoints = foreach($C in $DPSP.Children) {if($C.Children[1].text){$C.Children[1].text}}

    $MainOut | Add-Member -NotePropertyName "SupersedanceDependency" -NotePropertyValue $SupersedanceDependency

    $Distribution = New-Object -TypeName psobject
    $Distribution | Add-Member -NotePropertyName "DistributionPointGroups" -NotePropertyValue $DistributionGroups
    $Distribution | Add-Member -NotePropertyName "DistributionPoints" -NotePropertyValue $DistributionPoints
    $MainOut | Add-Member -NotePropertyName "Distrubution" -NotePropertyValue $Distribution

    
    $Collections = foreach($C0 in $CollectionMasterSP.Children){
        $CollectionOut = New-Object -TypeName psobject
        foreach($C in $C0.Children[1].Children){$CollectionOut | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}
        $DeploymentOut = New-Object -TypeName psobject
        foreach($C in $C0.Children[3].Children){$DeploymentOut | Add-Member -NotePropertyName $C.Children[1].tag -NotePropertyValue $C.Children[1].text}
        $CollectionOut | Add-Member -NotePropertyName "Deployment" -NotePropertyValue $DeploymentOut -Force
        $CollectionOut
    }

    $MainOut | Add-Member -NotePropertyName "Collection" -NotePropertyValue $Collections -Force

    if($Pass){$MainOut | StringConvert}
    else{$MainOut | StringConvert | ConvertTo-Json -Depth 99 | New-Item -Path $path -ItemType File -force}
    
}
Function UpdateHash {
    param(
        $Hash,
        $Pannel
    )
    if($Pannel){
        foreach($C in $Pannel.Children){
            if($Hash.$($C.Children[1].Tag).Type -eq 'Textbox'){
                $Hash.$($C.Children[1].Tag).Content = $C.Children[1].Text
            }
            elseif($Hash.$($C.Children[1].Tag).Type -eq 'ComboBox'){
                $Hash.$($C.Children[1].Tag).Index = $C.Children[1].SelectedIndex
                $Hash.$($C.Children[1].Tag).Content = $C.Children[1].SelectedItem.Content
            }
        }
    }
    else{
        Write-Host "No Panel" -ForegroundColor Yellow
    }
    $Hash
}
Function Start-Up {
    param([switch]$Clear)

    if($Clear){
        $ApplicationSP.Children.RemoveRange(0,$($ApplicationSP.Children.Count))
        $DeploymentTypeDynamicSP.Children.RemoveRange(0,$($DeploymentTypeDynamicSP.Children.Count))
        $DetectionMethodWISP.Children.RemoveRange(0,$($DetectionMethodWISP.Children.Count))
        $DetectionMethodFSSP.Children.RemoveRange(0,$($DetectionMethodFSSP.Children.Count))
        $DetectionMethodREGSP.Children.RemoveRange(0,$($DetectionMethodREGSP.Children.Count))
        $MasterDMSP.Children.RemoveRange(1,$($MasterDMSP.Children.Count))
        $DPGSP.Children.RemoveRange(1,$($DPGSP.Children.Count -2))
        $DPGSP.Children[0].Children[1].Text = $null
        $DPSP.Children.RemoveRange(1,$($DPSP.Children.Count -2))
        $DPSP.Children[0].Children[1].Text = $null
        $CollectionMasterSP.Children.RemoveRange(0,$($CollectionMasterSP.Children.Count))
        $TecnologyCB.SelectedIndex = -1
        $DetectionTypeCB.SelectedIndex = -1
    }

    $DetectionMethodWISP.Visibility = 2
    $DetectionMethodFSSP.Visibility = 2
    $DetectionMethodREGSP.Visibility = 2

    Add-SP -Mother $ApplicationSP -Hash $ApplicationHash
    Add-SP -Mother $DetectionMethodWISP -Hash $DetectionMethodWI
    Add-SP -Mother $DetectionMethodFSSP -Hash $DetectionMethodFS
    Add-SP -Mother $DetectionMethodREGSP -Hash $DetectionMethodREG
    Add-SP -Mother $DeploymentTypeDynamicSP -Hash $DeploymetTypeHash

    #Add-SP -Mother $DetectionMethodREGSP -Hash $DetectionMethodREG
    New-CollectionDeploymentSP -HashTemplateC $CollectionHASHTemplate -HashTemplateD $DeploymentHASHTemplate -Mother $CollectionMasterSP | Out-Null
    $TecnologyCB.SelectedIndex = 0
    $DetectionTypeCB.SelectedIndex = 0
}
Function Call-Buttons {
    Param($Buttons)

$ButtonWindow = @"
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Rory's SSCCM Importer" WindowStartupLocation = "CenterScreen"
        SizeToContent = "WidthAndHeight" 
        ShowInTaskbar = "false" 
        Background = "White" 
        ResizeMode = "NoResize"
    >
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight">
    <StackPanel/>
"@
    Add-Type -AssemblyName PresentationFramework
    $ButtonReader = (New-Object System.Xml.XmlNodeReader $xaml)
    $ButtonWindow = [Windows.Markup.XamlReader]::Load($reader)

    $async = $window.Dispatcher.InvokeAsync({
        $window.Showdialog() | Out-Null
    })

    $async.Wait() | Out-Null
}
Function ImportSCCM {
    param(
        [Switch]$ApplicationDOIT,
        [Switch]$DeploymentTypeDOIT,
        [Switch]$SuperDependDOIT,
        [Switch]$DistributionDOIT,
        [Switch]$CollectionDOIT,
        [Switch]$DeploymentDOIT,
        [Switch]$ADGroupDOIT
    )

    $Template = Export-Template -Pass

    #$ApplicationDOIT = $Template.AddApplication
    #$DeploymentTypeDOIT = $Template.AddDeploymentType
    #$DistributionDOIT = $Template.DistributeSource
    #$CollectionDOIT = $Template.AddCollection
    #$DeploymentDOIT = $Template.AddDeployment
    #$ADGroupDOIT = $Template.AddAdGroup
    
    $Application = $null
    $DeploymentType = $null
    if(Connect-RCM -Silent){
        if($ApplicationDOIT){
            $Application = New-RCMAppFromTemplate -template $Template
            $GUIBIT = $ApplicationSP.children | ?{$_.Children[1].tag -eq "ApplicationName"} | %{$_.Children[1]}
            IF($GUIBIT.text -ne $Application.LocalizedDisplayName){$GUIBIT.text = $Application.LocalizedDisplayName}
        }
        else{
            $Application = Get-RCMApp -Name $Template.Name
        }
        if($DeploymentTypeDOIT -and $Application){
            Check-App -Template $Template
            $DeploymentType = New-RCMDeploymentTypeFromTemplate -Template $Template -Application $Application
            $GUIBIT = $DeploymentTypeSP.children | ?{$_.Children[1].tag -eq "DeploymentTypeName"} | %{$_.Children[1]}
            IF(($GUIBIT.text -ne $DeploymentType.LocalizedDisplayName)-and($GUIBIT.text -ne $ApplicationSP.LocalizedDisplayName)){$GUIBIT.text = $DeploymentType.LocalizedDisplayName}
        }
        else{
            $DeploymentType = Get-RCMDeploymentType -ApplicationName $Template.Name
        }
        if($SuperDependDOIT -and $Application -and $DeploymentType){
            New-RCMSuperDependFromTemplate -Template $Template -Application $Application -DeploymentType $DeploymentType
        }
        if($Application -and $DeploymentType -and $DistributionDOIT){
            $DistrubutionSplat = $Template.Distrubution | ConvertTo-Hashtable -RemoveNull
            $Distribution = Add-RCMDistribution -APPName $Application.LocalizedDisplayName @DistrubutionSplat
        }
        if($CollectionDOIT){
            if($ADGroupDOIT){$collection = New-RCMCollectionFromTemplate -Template $Template -Application $Application -Group}
            else{$collection = New-RCMCollectionFromTemplate -Template $Template -Application $Application}
            #foreach(){}
        }
        else{
            $CollectionName = $Template.Collection | %{$_.name} 
            if(!$CollectionName){$CollectionName = $Template.Name}
            $collection = $CollectionName | %{Get-RCMCollection -Name $_}
        }
        if($DeploymentDOIT){
            if($Application){$ApplicationName = $Application.LocalizedDisplayName}
            else{$ApplicationName = $Template.Name}
            if($Application){
                New-RCMDeploymentFromTemplate -ApplicationName $ApplicationName -Template $Template
            }
        }
    }
    else{
        Write-Host "No SCCM Connection"
    }
}
Function Add-DetectionMethod{

    $DetectionSP = New-Object System.Windows.Controls.StackPanel
    
    $MP = New-Object System.Windows.Controls.StackPanel
    $MP.Orientation = 'Horizontal'
    $MP.Tag = 'DetectionMethodConnector'
    $LB= New-Object System.Windows.Controls.Label
    $LB.HorizontalContentAlignment = "Right"
    $LB.Width = 120
    $CB = New-Object System.Windows.Controls.ComboBox
    $CB.Width = 50
    $CB.Tag = "JoinBox"
    $CB.StaysOpenOnEdit = $true
    foreach($C in "and","or"){
        $N = New-Object System.Windows.Controls.ComboBoxItem
        $N.Content = $C
        $CB.Items.Add($N) | Out-Null
    }
    $CB.SelectedIndex = 0
    $BT = New-Object System.Windows.Controls.Button
    $BT.Content = "  -  "
    $BT.ToolTip = "Remove Below"
    $BT.Add_Click({
        $this.Parent.Parent.Parent.Children.Remove($this.Parent.Parent)
    })
    $MP.AddChild($LB)
    $MP.AddChild($CB)
    $MP.AddChild($BT)
    $DetectionSP.AddChild($MP)

    $ComboboxObject = New-ComboBoxPanel -LabelName 'Type:' -LogicalName "Type" -CBItems "WindowsInstaller","FileSystem","Registry" -GRandMother $DetectionSP -Passthrough -Index 1
    $ComboboxObject.Parent.Tag = "TYPE"
    $MsiSP = New-Object System.Windows.Controls.StackPanel
    $MsiSP.Tag = "MSI"
    Add-SP -Mother $MsiSP -Hash $DetectionMethodWI 
    $DetectionSP.AddChild($MsiSP)
    $FSSP = New-Object System.Windows.Controls.StackPanel
    $FSSP.Tag = "FILE"
    Add-SP -Mother $FSSP -Hash $DetectionMethodFS
    $DetectionSP.AddChild($FSSP)
    $RegSP = New-Object System.Windows.Controls.StackPanel
    $RegSP.Tag = "REG"
    Add-SP -Mother $RegSP -Hash $DetectionMethodREG
    $DetectionSP.AddChild($RegSP)

    #$this.Parent.Parent
    $ComboboxObject.add_SelectionChanged({
        switch ($this.SelectedItem.Content){
            "WindowsInstaller" {
                $this.Parent.Parent.children[2].Visibility = 0
                $this.Parent.Parent.children[3].Visibility = 2
                $this.Parent.Parent.children[4].Visibility = 2
            }
            "FileSystem" {
                $this.Parent.Parent.children[2].Visibility = 2
                $this.Parent.Parent.children[3].Visibility = 0
                $this.Parent.Parent.children[4].Visibility = 2
            }
            "Registry" {
                $this.Parent.Parent.children[2].Visibility = 2
                $this.Parent.Parent.children[3].Visibility = 2
                $this.Parent.Parent.children[4].Visibility = 0
            }
            default    {
                $this.Parent.Parent.children[2].Visibility = 2
                $this.Parent.Parent.children[3].Visibility = 2
                $this.Parent.Parent.children[4].Visibility = 2
            }
        }
    })
    $ComboboxObject.SelectedIndex = 0

    $MasterDMSP.AddChild($DetectionSP)
}
function Check-Template{
    $Template = Export-Template -Pass
    $Template = Get-Content -Path "C:\Scripts\TestTemplate.json" -raw | ConvertFrom-Json
}
Function Check-App{
    Param(
        $Template=$(Export-Template -Pass),
        [switch]$TeplateOnly
    )
    $AppIsGood = $true
    [array]$Message = @()
    if($Template.Name){
        [array]$FullList = Get-RCMApp -Name * -Fast | %{$_.LocalizedDisplayName}
        if($FullList -contains $Template.Name){
            $Message += "Applicaiton Exists in SCCM"
            $AppIsGood = $false
            $GuiObject = $ApplicationSP.Children | ?{$_.Children[1].tag -eq "ApplicationName"}
            $GuiObject.Background = "tomato"
        }
        else{
            $GuiObject = $ApplicationSP.Children | ?{$_.Children[1].tag -eq "ApplicationName"}
            $GuiObject.Background = "white"
        }
    }
    else{
        $Message += "Application Name is required"
        $AppIsGood = $false
        $GuiObject = $ApplicationSP.Children | ?{$_.Children[1].tag -eq "ApplicationName"}
        $GuiObject.Background = "tomato"
    }

    $GuiObject = $ApplicationSP.Children | ?{$_.Children[1].tag -eq "SCCMFolder"}
    if($TeplateOnly -or !(Test-RCMPath $Template.App.SCCMFolder)){
        $Message += "SCCM Path does not exist"
        $AppIsGood = $false
        $GuiObject.Background = "tomato"
    }
    else{
        $GuiObject.Background = "white"
    }
    $AppIsGood
    $Message | Write-Host
}
Function Check-DeploymentType {
    Param(
        $Template=$(Export-Template -Pass),
        [switch]$TeplateOnly
    )
    $IsGood = $true
    [array]$Message = @()
    $GuiObject = $DeploymentTypeDynamicSP.Children | ?{$_.Children[1].tag -eq "ContentLocation"}
    if($Template.DeploymentType.ContentLocation){
        if($GuiObject){$GuiObject.Background = "white"}
    }
    else{
        $Message += "ContentLocation is needed"
        $IsGood = $false
        if($GuiObject){$GuiObject.Background = "tomato"}
    }
    if($Template.DeploymentType.Tecnology -eq "Script"){
        $GuiObject = $DeploymentTypeDynamicSP.Children | ?{$_.Children[1].tag -eq "InstallScript"}
        if($Template.DeploymentType.InstallScript){
            if($GuiObject){$GuiObject.Background = "white"}
        }
        else{
            $Message += "Install command is needed"
            $IsGood = $false
            if($GuiObject){$GuiObject.Background = "tomato"}
        }
    }
    elseif($Template.DeploymentType.Tecnology -eq "MSI"){
        $GuiObject = $DeploymentTypeDynamicSP.Children | ?{$_.Children[1].tag -eq "MsiName"}
        if($Template.DeploymentType.MsiName){
            if($GuiObject){$GuiObject.Background = "white"}
        }
        else{
            $Message += "Install command is needed"
            $IsGood = $false
            if($GuiObject){$GuiObject.Background = "tomato"}
        }
    }
    elseif($Template.DeploymentType.Tecnology -eq "appv"){
        $GuiObject = $DeploymentTypeDynamicSP.Children | ?{$_.Children[1].tag -eq "msi"}
        if($Template.DeploymentType.AppVname){
            if($GuiObject){$GuiObject.Background = "white"}
        }
        else{
            $Message += "Install command is needed"
            $IsGood = $false
            if($GuiObject){$GuiObject.Background = "tomato"}
        }
    }
    $IsGood
}
Function Check-Distribution {
    Param(
        $Template=$(Export-Template -Pass)
    )
    if($RCMisConnected -or (Connect-RCM -Silent)){
        
    }
    else{
        $false
        return
    }
    if($Template.Distrubution.DistributionPointGroups -or $Template.Distrubution.DistributionPoints){
        $Truith = $true
        #$DistributionGroups = foreach($C in $DPGSP.Children){if($C.Children[1].text){$C.Children[1].text}}
        #$DistributionPoints = foreach($C in $DPSP.Children) {if($C.Children[1].text){$C.Children[1].text}}
        $GuiObjects = @()
        
        $DPGSP.Children | %{$_.Background = "white"}
        $DPSP  | %{$_.Background = "white"}

        if(!$DistributionPointGroups){
            [array]$DistributionPointGroups = get-RCMDistributionPointGroup | %{$_.name}
        }
        foreach($DG in $Template.Distrubution.DistributionPointGroups){
            if($DistributionPointGroups -notcontains $DG){
                $Truith = $false
                $GuiObjects += $DPGSP.Children | ?{$_.Children[1].text -eq $DG}
            }
        }
        if(!$DistributionPoints){
            [array]$DistributionPoints = get-RCMDistributionPoint | %{if($_.ItemName -match '\[\"Display=[\\]*([^\"]*?)[\\]*\"\]'){$Matches[1]}}
        }
        
        foreach($DP in $Template.Distrubution.DistributionPoints){
            if($DistributionPoints -notcontains $DP){
                $Truith = $false
                $GuiObjects += $DPSP.Children | ?{$_.Children[1].text -eq $DP}
            }
        }
        $GuiObjects | %{$_.Background = "tomato"}
        $Truith
        return
    }
    else{
        $false
        return
    }
}
Function Set-RComboBoxOptions {
    Param([array]$ItemList,$ComboBox,$ReturnAction=$false,$FilterLogic)
    $Tag = New-Object -TypeName psobject
    $Tag | Add-Member -NotePropertyName "ReturnAction" -NotePropertyValue $ReturnAction
    if($FilterLogic){
        $Tag | Add-Member -NotePropertyName "FilterLogic" -NotePropertyValue $FilterLogic
    }
    
    $ComboBox.tag = $Tag
    if($ComboBox.ItemsSource){
        $ComboBox.ItemsSource = $null
    }
    if($ComboBox.Items){
        $($ComboBox.Items.Count -1)..0 | %{$ComboBox.Items.RemoveAt($_)}
    }
    $ComboBox.StaysOpenOnEdit = $true
    $ItemList | %{New-Object System.Windows.Controls.ComboBoxItem -Property @{Content=$_}} | %{$ComboBox.Items.Add($_)} | Out-Null
    #$ComboBox.add_GotFocus({if(!$this.IsDropDownOpen){$this.IsDropDownOpen = $true}})
    $ComboBox.add_Keydown({
        if(!$this.text -and !$this.IsDropDownOpen){$this.IsDropDownOpen = $true}
    })
    $ComboBox.add_Keyup({
        $ExcludedKeys = 'Up','Right','Left','Down','Return','LeftCtrl','RightCtrl'
        if($ExcludedKeys -notcontains $_.key){
            $thisCB = $this
            if($this.tag.FilterLogic){
                $thisCB.Items.Filter = {(($args.content -like "*$($thisCB.text)*")-or($args.content -eq $thisCB.SelectedItem.Content)) -and ($(&$thisCB.tag.FilterLogic) -notcontains $args.content)}
            }
            else{
                $thisCB.Items.Filter = {(($args.content -like "*$($thisCB.text)*")-or($args.content -eq $thisCB.SelectedItem.Content))}
            }
        }
        elseif ($_.Key -eq 'RETURN' -and $this.tag.ReturnAction){
            &$this.tag.ReturnAction
        }
        elseif($_.key -eq 'Down'){
            $this.IsDropDownOpen = $true
        }
        if(!$this.text -and !$this.IsDropDownOpen){$this.IsDropDownOpen = $true}
    })
    if($FilterLogic){
        $ComboBox.add_DropDownOpened({
            $thisCB = $this
            $this.Items.Filter = {(($args.content -like "*$($thisCB.text)*")-or($args.content -eq $thisCB.SelectedItem.Content)) -and ($(&$thisCB.tag.FilterLogic) -notcontains $args.content)}
            #Write-Host "$(&$this.tag.FilterLogic)"
        })
    }
}
#endregion

Function Call-RCMGui {
param(
    [switch]$DirectImport
)
$ConnectedRCM = Connect-RCM -Silent
#region GUI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Rory's SCCM Importer" WindowStartupLocation = "CenterScreen"
    SizeToContent = "WidthAndHeight" 
    ShowInTaskbar = "True" 
    Background = "White" 
    ResizeMode = "NoResize"
>
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight">
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight">
            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="PIC">
                <Image Source = "$PSScriptRoot\Logo78b.png" Stretch = "Fill"/>
            </StackPanel>
            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="ButtonsSP">
                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                    <Button x:Name = "ImportFromFileBT" Height = "26" Content = '  Import From File  ' ToolTip = "Browse" />
                    <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="SCCMSearchBoxCB" Height = "26" Width = "380" KeyboardNavigation.TabIndex="0"/>
                    <Button x:Name = "ImportFromSCCMBT" Height = "26" Content = '  Create Template From SCCM ' ToolTip = "Create from selected application" />
                    <Button x:Name = "ExportTemplate" Height = "26" Content = '  Export Template ' ToolTip = "Export" />
                </StackPanel>
                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                    <Button x:Name = "ImportFromSource" Height = "26" Content = '  Fill In other information from source files ' ToolTip = "Don't trust this to much it's a guess" />
                    <Button x:Name = "ClearGUI" Height = "26" Content = '  Clear GUI ' ToolTip = "Clear" />
                    <Button x:Name = "CreateImportBundle" Height = "26" Content = '  Create Import Bundle ' ToolTip = "Bundles are fun" />
                </StackPanel>
                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                    <Button x:Name = "AddToSCCM" Height = "26" Content = '  Add To SCCM ' ToolTip = "DO IT!" />
                    <CheckBox Margin="5" Name="ApplicationCH" Content = "Applicaiton" IsChecked="true"/>
                    <CheckBox Margin="5" Name="DeploymentTypeCH" Content="DeploymentType" IsChecked="true"/>
                    <CheckBox Margin="5" Name="SupDependencyCH" Content="Supersede / Depend" IsChecked="true"/>
                    <CheckBox Margin="5" Name="DistributionCH" Content="Distribution" IsChecked="true"/>
                    <CheckBox Margin="5" Name="CollectionCH" Content="Collection" IsChecked="true"/>
                    <CheckBox Margin="5" Name="DeploymentCH" Content="Deployment" IsChecked="true"/>
                    <CheckBox Margin="5" Name="ADGroupsCH" Content="AD Group(s)" IsChecked="true"/>
                </StackPanel>
            </StackPanel>
        </StackPanel>
        <TabControl Name="TabC" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <TabItem Header="Application">
                <ScrollViewer HorizontalScrollBarVisibility="Auto">
                    <StackPanel Orientation = "Vertical">
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Application" Width = "120"/>
                            <Button x:Name = "AddApplication" Height = "26" Content = " Add Application to SCCM " ToolTip = "Application" Width = "380" />
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="ApplicationSP"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Deployment Type">
                <ScrollViewer HorizontalScrollBarVisibility="Auto" Name = "DeploymentTypeScrollViewer">
                    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DeploymentTypeSP">
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "DeploymentType" Width = "120"/>
                            <Button x:Name = "AddDeploymentType" Height = "26" Content = " Add Deployment Type To Application " ToolTip = "Deployment Type" Width = "380" />
                        </StackPanel>
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Tecnoligy:" Width = "120" HorizontalContentAlignment="Right"/>
                            <ComboBox Name="TecnologyCB" Width='600'>
                                <ComboBoxItem>MSI</ComboBoxItem>
                                <ComboBoxItem>Script</ComboBoxItem>
                                <ComboBoxItem>App-V</ComboBoxItem>
                            </ComboBox>
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DeploymentTypeDynamicSP"/>
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Detection Method" Width = "120" HorizontalContentAlignment="Right"/>
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name = "MasterDMSP">                        
                            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" >
                                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" Tag = "TYPE">
                                    <Label Height = "26" Content = "Type:" Width = "120" HorizontalContentAlignment="Right"/>
                                    <ComboBox Name="DetectionTypeCB" Width='600'>
                                        <ComboBoxItem>WindowsInstaller</ComboBoxItem>
                                        <ComboBoxItem>FileSystem</ComboBoxItem>
                                        <ComboBoxItem>Registry</ComboBoxItem>
                                        <ComboBoxItem>App-V / will do later</ComboBoxItem>
                                    </ComboBox>
                                </StackPanel>
                                <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DetectionMethodWISP" Tag = "MSI"/>
                                <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DetectionMethodFSSP" Tag = "FILE"/> 
                                <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DetectionMethodREGSP" Tag = "REG"/>
                            </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "" Width = "120" HorizontalContentAlignment="Right"/>
                            <Button x:Name = "AddDetectionMethodBT" Height = "26" Content = "  +  " ToolTip = "Add Detection Method" />
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Supersedence / Dependency">
                <ScrollViewer HorizontalScrollBarVisibility="Auto">
                    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight">
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "" Width = "120"/>
                            <Button x:Name = "SupersedenceButton" Height = "26" Content = " add Supersedence / Dependency " ToolTip = "add Supersedence / Dependency" Width = "380" />
                        </StackPanel>
                        
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Supersedence"/>
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="SupersedenceSP">
                            <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                <Label Height = "26" Content = "Supersede:" Width = "150" HorizontalContentAlignment="Right"/>
                                <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="SupersedenceCB" Height = "26" Width = "470" KeyboardNavigation.TabIndex="0"/>
                                <CheckBox Name="UninstallCB" Content="Uninstall" IsChecked="true"/>
                            </StackPanel>
                            <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                <Label Height = "26" Content = " " Width = "150" HorizontalContentAlignment="Right"/>
                                <Button x:Name = "AddSupersedenceCBBT" Height = "26" Content = "  +  " ToolTip = "Add Supersedence" />
                            </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Dependency"/>
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DependencySP">
                            <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                <Label Height = "26" Content = "Dependency:" Width = "150" HorizontalContentAlignment="Right"/>
                                <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="DependencyCB" Height = "26" Width = "470" KeyboardNavigation.TabIndex="0"/>
                                <CheckBox Name="InstallCB" Content="Install" IsChecked="true"/>
                            </StackPanel>
                            <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                <Label Height = "26" Content = " " Width = "150" HorizontalContentAlignment="Right"/>
                                <Button x:Name = "AddDependencyCBBT" Height = "26" Content = "  +  " ToolTip = "Add Dependency" />
                            </StackPanel>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Distribution">
                <ScrollViewer HorizontalScrollBarVisibility="Auto">
                    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight">
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Application" Width = "120"/>
                            <Button x:Name = "AddDistribution" Height = "26" Content = " Distribute content " ToolTip = "Distribute content" Width = "380" />
                        </StackPanel>
                        <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DistributionSP">
                            <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                <Label Height = "26" Content = "Distrubution"/>
                            </StackPanel>
                            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DPGSP">
                                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                    <Label Height = "26" Content = "Deployment Point Group:" Width = "150" HorizontalContentAlignment="Right"/>
                                    <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="DistributionPointGroupCB" Height = "26" Width = "570" KeyboardNavigation.TabIndex="0"/>
                                </StackPanel>
                                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                    <Label Height = "26" Content = " " Width = "150" HorizontalContentAlignment="Right"/>
                                    <Button x:Name = "AddDeploymentPointGroupBT" Height = "26" Content = "  +  " ToolTip = "Add DeploymentPointGroup" />
                                </StackPanel>
                            </StackPanel>
                            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="DPSP">
                                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                    <Label Height = "26" Content = "Deployment Point:" Width = "150" HorizontalContentAlignment="Right"/>
                                    <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="DistributionPointCB" Height = "26" Width = "570" KeyboardNavigation.TabIndex="0"/>
                                </StackPanel>
                                <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                                    <Label Height = "26" Content = " " Width = "150" HorizontalContentAlignment="Right"/>
                                    <Button x:Name = "AddDeploymentPointBT" Height = "26" Content = "  +  " ToolTip = "Add DeploymentPointGroup" />
                                </StackPanel>
                            </StackPanel>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            <TabItem Header="Collection / Deployment">
                <ScrollViewer HorizontalScrollBarVisibility="Auto" Name="CollectionScrollViewer">
                    <StackPanel Orientation = "Vertical" >
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Collection - Deployment"/>
                            <Button x:Name = "CollectionDeploymentBT" Height = "26" Content = " create Collection(s) " ToolTip = "Application" Width = "280" />
                            <Button x:Name = "DeploymentBT" Height = "26" Content = " create Deployment(s) " ToolTip = "Application" Width = "280" />
                        </StackPanel>
                            <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="CollectionMasterSP">
                        </StackPanel>
                        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
                            <Label Height = "26" Content = "Add - New Collection - Deployment"/>
                            <Button x:Name = "NewCollection" Height = "26" Content = "  +  " ToolTip = "Add DeploymentPointGroup" />
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </StackPanel>
</Window>
"@

Add-Type -AssemblyName PresentationFramework
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

#endregion
#$ApplicationHash.SCCMFolder.BrowseObject
#Invoke-RWPFTreeBrowse -Root $ApplicationHash.SCCMFolder.Root -Object $ApplicationHash.SCCMFolder.BrowseObject
#region HASHTabels ysed to somplify changes needed to be made
#try{
#    $ADTree = Get-RADtree
#    $ADRoot = $($ADTree[0].FullName -split "\\")[0]
#}
#catch{
#    $ADTree = @()
#   $ADRoot = "AD"
#}
Write-Host "R1 : $("$RCMSiteCode\Application")" -ForegroundColor Yellow
$ApplicationHash = [ordered]@{
    "ApplicationName" = @{LabelName="Application Name:"; LogicalName="ApplicationName"; Type="Textbox"; Content=""   }
    "SCCMFolder" =      @{LabelName="SCCMFolder:";       LogicalName="SCCMFolder";      Type="TextboxBrowse"; Content="" ;Root="$RCMSiteCode\Application";ObjectBlock={Get-RCMChildFolder -Root -Recurse -FolderType Application}} #;BrowseObject=$(Get-RCMChildFolder -Root -Recurse -FolderType Application)
    "IconFile" =        @{LabelName="IconFile:";         LogicalName="IconFile";        Type="Textbox"; Content=""   }
    "Comment" =         @{LabelName="Comment:";          LogicalName="Comment";         Type="Textbox"; Content=""   }
    "LocalizedName" =   @{LabelName="Localized Name:";   LogicalName="LocalizedName";   Type="Textbox"; Content=""   }
    "Publisher" =       @{LabelName="Publisher:";        LogicalName="Publisher";       Type="Textbox"; Content=""   }
    "Version" =         @{LabelName="Version:";          LogicalName="Version";         Type="Textbox"; Content=""   }
    "AllowTaskSequence"=@{LabelName="Allow Task Sequence:"; LogicalName="AllowTaskSequence";Type="ComboBox";CBItems=@($true,$false);Index=1;Content=""}
}
$DeploymetTypeHash = [ordered]@{
    "CopyFrom" =                  @{LabelName="Copy From:";            LogicalName="CopyFrom";                 In=@("msi","script","App-V"); Type="Textbox"; Content=""   }
    "DeploymentTypeName" =        @{LabelName="DT Name:";              LogicalName="DeploymentTypeName";       In=@("msi","script","App-V"); Type="Textbox"; Content=""   }
    "ContentLocation" =           @{LabelName="Content Location:";     LogicalName="ContentLocation";          In=@("msi","script","App-V"); Type="Textbox"; Content=""   }
    "ContentFallback" =           @{LabelName="Content Fallback:";     LogicalName="ContentFallback";          In=@("msi","script","App-V"); Type="ComboBox";CBItems=@($true,$false);Index=1;Content=""}
    "SlowNetworkDeploymentMode" = @{LabelName="Slow Network Mode:";    LogicalName="SlowNetworkDeploymentMode";In=@("msi","script","App-V"); Type="ComboBox";CBItems=@("Download","DoNothing");Index=0;Content=""}
    "InstallationBehaviorType" =  @{LabelName="Installation Behavior:";LogicalName="InstallationBehaviorType"; In=@("msi","script","App-V"); Type="ComboBox";CBItems=@("InstallForSystem","InstallForUser","InstallForSystemIfResourceIsDeviceOtherwiseInstallForUser");Index=0;Content=""}
    "UserInteractionMode" =       @{LabelName="User Interaction:";     LogicalName="UserInteractionMode";      In=@("msi","script","App-V"); Type="ComboBox";CBItems=@("Hidden","Normal","Maximized","Minimized");Index=0;Content=""}
    "LogonRequirementType" =      @{LabelName="Logon Required:";       LogicalName="LogonRequirementType";     In=@("msi","script","App-V"); Type="ComboBox";CBItems=@("WhetherOrNotUserLoggedOn","OnlyWhenUserLoggedOn");Index=0;Content=""}
    "InstallScript" =             @{LabelName="Install Script:";       LogicalName="InstallScript";            In=@("script");               Type="Textbox"; Content=""   }
    "UninstallScript" =           @{LabelName="Uninstall Script:";     LogicalName="UninstallScript";          In=@("script");               Type="Textbox"; Content=""   }
    "RepairCommand" =             @{LabelName="Repair Command:";       LogicalName="RepairCommand";            In=@("script");               Type="Textbox"; Content=""   }
    "MsiName" =                   @{LabelName="Msi Name:";             LogicalName="MsiName";                  In=@("msi")                 ; Type="Textbox"; Content=""   }
    "MsTName" =                   @{LabelName="Mst Name:";             LogicalName="MstName";                  In=@("msi")                 ; Type="Textbox"; Content=""   }
    "MaximumRuntimeMins" =        @{LabelName="Maximum Runtime:";      LogicalName="MaximumRuntimeMins";       In=@("msi","script","App-V"); Type="Textbox"; Content="120"}
    "EstimatedRuntimeMins" =      @{LabelName="Estimated Runtime:";    LogicalName="EstimatedRuntimeMins";     In=@("msi","script","App-V"); Type="Textbox"; Content="5"  }
    "CloseExecutables" =          @{LabelName="Close Executables:";    LogicalName="CloseExecutables";         In=@("msi","script","App-V"); Type="Textbox"; Content=""   }
}
$DetectionMethodWI = [ordered]@{
    "ProductCode" =               @{LabelName="Product Code:";         LogicalName="ProductCode";              Type="Textbox"; Content=""   }
    "ExpectedValue" =             @{LabelName="Expected Version:";     LogicalName="ExpectedValue";            Type="Textbox"; Content=""   }
    "Operator" =                  @{LabelName="Operator:";             LogicalName="Operator";                 Type="ComboBox";CBItems=@('Equals',"Not Equal to",'Greater than or equal to','Grater Than','Less Than','Less Than or equal to');Index=0;Content=""}
}
$DetectionMethodREG = [ordered]@{
    'RegistryHive' =              @{LabelName="Registry Hive:";        LogicalName="RegistryHive";             Type="ComboBox";CBItems=@('HKEY_CLASSES_ROOT','HKEY_CURRENT_USER','HKEY_LOCAL_MACHINE','HKEY_USERS','HKEY_CURRENT_CONFIG');Index=2;Content=""}
    "KeyPath" =                   @{LabelName="Key Path:";             LogicalName="KeyPath";                  Type="Textbox"; Content=""   }
    "RegistryValueName" =         @{LabelName="Registry Value Name:";  LogicalName="RegistryValueName";        Type="Textbox"; Content=""   }
    "RegistryValueType" =         @{LabelName="Registry Value Type:";  LogicalName="RegistryValueType";        Type="ComboBox";CBItems=@('String','Integer','Version');Index=0;Content=""}
    "RegistryValue" =             @{LabelName="Registry Value:";       LogicalName="RegistryValue";            Type="Textbox"; Content=""   }
    "Operator" =                  @{LabelName="Operator:";             LogicalName="Operator";                 Type="ComboBox";CBItems=@('Equals','Not equal to','Greater than','Less than','Begins with','Does not begin with','Ends with','Does not end with','Contains','Does not contain','Between','One of','None of','Greater than or equal to','Less than or equal to');Index=0;Content=""}
}
$DetectionMethodFS = [ordered]@{
    "Type" =                      @{LabelName="Type:";                 LogicalName="Type";                     Type="ComboBox";CBItems=@('File','Folder');Index=0;Content=""}
    "Path" =                      @{LabelName="Path:";                 LogicalName="Path";                     Type="Textbox"; Content=""   }
    "Property" =                  @{LabelName="Property:";             LogicalName="Property";                 Type="ComboBox";CBItems=@('Existence','Date Modified','Date Created','Version','Size (Bytes)');Index=0;Content=""}
    "Operator" =                  @{LabelName="Operator:";             LogicalName="Operator";                 Type="ComboBox";CBItems=@('Equals',"Not Equal to",'Greater than or equal to','Grater Than','Less Than','Less Than or equal to','Between','One of','None of');Index=0;Content=""}
    "Value" =                     @{LabelName="Value:";                LogicalName="Value";                    Type="Textbox"; Content=""   }
}
$DetectionMethodCleanHash = @{
    'WindowsInstaller'= 'Path','RegistryHive','RegistryPath','RegistryKey','RegistryValue','RegistryValueType' #'ProductCode','Version','Operator'
    'FileSystem'=       'ProductCode','RegistryHive','RegistryPath','RegistryKey','RegistryValue','RegistryValueType' #'Path','Version','Operator'
    'Registry' =        'ProductCode','Path','VersionNumber'#'RegistryHive','RegistryPath','RegistryKey','RegistryValue','RegistryValueType','Operator'
}
$CollectionHASHTemplate = [ordered]@{
    "Name" =                   @{LabelName="Collection Name:";     LogicalName="Name";                   Type="Textbox"; Content=""}
    "LimitingCollectionName" = @{LabelName="Limiting Collection:"; LogicalName="LimitingCollectionName"; Type="Textbox"; Content=""}
    "Type" =                   @{LabelName="Type:";                LogicalName="Type";                   Type="ComboBox";CBItems=@("UserCollection","DeviceCollection");Index=1;Content=""}
    "Folder" =                 @{LabelName="SCCM Folder:";         LogicalName="Folder";                 Type="TextboxBrowse"; Content="" ;ObjectBlock={@(Get-RCMChildFolder -Recurse -FolderType Device_Collection -Root)+@($(Get-RCMChildFolder -Root -Recurse -FolderType User_Collection))};Root=$RCMSiteCode}
    #"Purpose" =                @{LabelName="Purpose:";             LogicalName="Purpose";                Type="ComboBox";CBItems=@("Install","Uninstall");Index=0;Content=""}
    "GroupName" =              @{LabelName="GroupName:";           LogicalName="GroupName";              Type="Textbox"; Content=""}
    "GroupOU" =                @{LabelName="GroupOU:";             LogicalName="GroupOU";                Type="TextboxBrowse"; Content="" ;ObjectBlock={Get-RADtree};Root=$false} #;ObjectBlock={Get-RADtree}
}
$DeploymentHASHTemplate = [ordered]@{
    "DeployAction" =           @{LabelName="Deploy Action:";       LogicalName="DeployAction";           Type="ComboBox";CBItems=@("Install","Uninstall");Index=0;Content=""}
    "DeployPurpose" =          @{LabelName="Deploy Purpose:";      LogicalName="DeployPurpose";          Type="ComboBox";CBItems=@("Required","Available");Index=0;Content=""}
    "AllowRepairApp" =         @{LabelName="Allow Users to repair Applicaion:";  LogicalName="AllowRepairApp"; Type="ComboBox";CBItems=@("true","false");Index=1;Content=""}
    "PreDeploy" =              @{LabelName="Pre Deploy to primary device:" ;     LogicalName="PreDeploy"     ; Type="ComboBox";CBItems=@("true","false");Index=1;Content=""}
    "CloseRunningExe" =        @{LabelName="Close Running Executables:" ;     LogicalName="CloseRunningExe"     ; Type="ComboBox";CBItems=@("true","false");Index=1;Content=""}
    "UserNotification" =       @{LabelName="User Notification:";   LogicalName="UserNotification";       Type="ComboBox";CBItems=@("DisplayAll","DisplaySoftwareCenterOnly","HideAll");Index=1;Content=""}
    "OverrideServiceWindow" =  @{LabelName="Install Outside Service Window:"; LogicalName="OverrideServiceWindow"; Type="ComboBox";CBItems=@("true","false");Index=0;Content=""}
    "RebootOutsideServiceWindow" = @{LabelName="Reboot Outside Service Window:";  LogicalName="RebootOutsideServiceWindow"; Type="ComboBox";CBItems=@("true","false");Index=1;Content=""}
}
#endregion
#region GuiObjectDeclirations
$AddApplication = $Window.FindName('AddApplication')
#$AddApplication = $Window.FindName('AddApplication')
$AddDeploymentPointBT = $Window.FindName('AddDeploymentPointBT')
$AddDeploymentPointGroupBT = $Window.FindName('AddDeploymentPointGroupBT')
#$AddDeploymentType = $Window.FindName('AddDeploymentType')
$AddDeploymentType = $Window.FindName('AddDeploymentType')
$AddDetectionMethodBT = $Window.FindName('AddDetectionMethodBT')
$AddDistribution = $Window.FindName('AddDistribution')
$AddToSCCM = $Window.FindName('AddToSCCM')
$ADGroupsCH  = $Window.FindName('ADGroupsCH')
$ApplicationCH  = $Window.FindName('ApplicationCH')
$ApplicationSP = $Window.FindName('ApplicationSP')
$ApplicationSP.Tag = "ApplicationHash"
$ClearGUI = $Window.FindName('ClearGUI')
$CollectionCH  = $Window.FindName('CollectionCH')
$CollectionDeployment = $Window.FindName('CollectionDeployment')
$CollectionDeploymentBT = $Window.FindName('CollectionDeploymentBT')
$CollectionMasterSP = $Window.FindName('CollectionMasterSP')
$CollectionScrollViewer = $Window.FindName('CollectionScrollViewer')
$CreateImportBundle = $Window.FindName('CreateImportBundle')
$DeploymentBT = $Window.FindName('DeploymentBT')
$DeploymentCH  = $Window.FindName('DeploymentCH')
$DeploymentTypeCH  = $Window.FindName('DeploymentTypeCH')
$DeploymentTypeDynamicSP = $Window.FindName('DeploymentTypeDynamicSP')
$DeploymentTypeDynamicSP.Tag = "DeploymetTypeHash"
$DeploymentTypeScrollViewer = $Window.FindName('DeploymentTypeScrollViewer')
$DeploymentTypeSP = $Window.FindName('DeploymentTypeSP')
$DeploymentTypeSP.Tag = "DeploymetTypeHash"
$DetectionGroupTypeCB = $Window.FindName('DetectionGroupTypeCB')
$DetectionMethodFSSP = $Window.FindName('DetectionMethodFSSP')
$DetectionMethodREGSP = $Window.FindName('DetectionMethodREGSP')
$DetectionMethodWISP = $Window.FindName('DetectionMethodWISP')
$DetectionTypeCB = $Window.FindName('DetectionTypeCB')
#$Distribution = $Window.FindName('Distribution')
$DistributionCH  = $Window.FindName('DistributionCH')
$DistributionPointCB = $Window.FindName('DistributionPointCB')
$DistributionPointGroupCB = $Window.FindName('DistributionPointGroupCB')
$DistributionSP = $Window.FindName('DistributionSP')
$DPGSP = $Window.FindName('DPGSP')
$DPSP = $Window.FindName('DPSP')
$ExportTemplate  = $Window.FindName('ExportTemplate')
$ImportFromFileBT = $Window.FindName('ImportFromFileBT')
$ImportFromSCCMBT = $Window.FindName('ImportFromSCCMBT')
$ImportFromSource  = $Window.FindName('ImportFromSource')
$MasterDMSP = $Window.FindName('MasterDMSP')
$NewCollection = $Window.FindName('NewCollection')
$Global:SCCMSearchBoxCB = $Window.FindName('SCCMSearchBoxCB')
$TabC = $Window.FindName('TabC')
$TecnologyCB = $Window.FindName('TecnologyCB')

$SupDependencyCH = $Window.FindName('SupDependencyCH')
$SupersedenceButton = $Window.FindName('SupersedenceButton')
$AddSupersedenceCBBT = $Window.FindName('AddSupersedenceCBBT')
$AddDependencyCBBT = $Window.FindName('AddDependencyCBBT')
$SupersedenceCB = $Window.FindName('SupersedenceCB')
$DependencyCB = $Window.FindName('DependencyCB')
$SupersedenceSP = $Window.FindName('SupersedenceSP')
$DependencySP = $Window.FindName('DependencySP')

#$DependencySP.gettype()
#$Window.gettype()

#region Events
$AddDeploymentPointBT.Add_Click({
    New-DeploymentItem -Mother $this.Parent.Parent -Type "DP"
})
$AddDeploymentPointGroupBT.Add_Click({
    New-DeploymentItem -Mother $this.Parent.Parent -Type "DPG"
})
$AddSupersedenceCBBT.Add_Click({
    New-DeploymentItem -Mother $this.Parent.Parent -Type "Supersedence" -Checkbox "Uninstall"
})
$AddDependencyCBBT.Add_Click({
    New-DeploymentItem -Mother $this.Parent.Parent -Type "Dependency" -Checkbox "Install"
})
$NewCollection.Add_Click({
    New-CollectionDeploymentSP -HashTemplateC $CollectionHASHTemplate -HashTemplateD $DeploymentHASHTemplate -Mother $CollectionMasterSP | Out-Null
    $CollectionScrollViewer.PageDown()
})
$TecnologyCB.add_SelectionChanged({
    $DeploymentTypeDynamicSP.Children | %{$_.children[1].tag}
    foreach($T in $DeploymentTypeDynamicSP.Children){
        $TAG = $T.Children[1].tag
        if($DeploymetTypeHash.$TAG.In -contains $this.SelectedItem.Content){
            $T.Visibility = 0
        }
        else{
             $T.Visibility = 2
        }
    }
    if($this.SelectedItem.Content -eq "App-V"){
        #$TecnologyCB.SelectedItem.Content
    }
    Check-DeploymentType
})
$ImportFromFileBT.Add_Click({
    $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $LastLocation
    $OpenFileDialog.Filter = "Json Template (*.json)|*.json"
    
    if($OpenFileDialog.ShowDialog()){
        $ImportFrom = $OpenFileDialog.FileName
        $Template = Import-JsonTemplate -Path $ImportFrom -Clear
    }
})

$IMPORT_SCCM_ACTION = {
    if($SCCMSearchBoxCB.text -notmatch "^\s*$"){
        $TEMPGUID = [guid]::NewGuid().guid
        $T = New-RCMTemplate -ApplicationName $SCCMSearchBoxCB.text
        $OutName = "NewAppTemplate-$RCMSiteCodeRaw-$TemplateAppName.json" -replace '\\|/|\:|\*|\"|<|>|\|',"_"
        $T | ConvertTo-Json -Depth 99 | New-Item -Path "$env:TEMP\$TEMPGUID.json" -ItemType file -Force

        $Template = Import-JsonTemplate -Path "$env:TEMP\$TEMPGUID.json" -Clear
    }
}

$ImportFromSCCMBT.Add_Click({&$IMPORT_SCCM_ACTION})
$ImportFromSource.Add_Click({
    $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $LastLocation

    if($OpenFileDialog.ShowDialog()){
        $ImportFrom = $OpenFileDialog.FileName
        $Base = Export-Template -Pass
        $NewTemplate = New-RcmTemplateFromSourceFile -Template $Base -MainInstaller $ImportFrom 
        Load-Template -Template $NewTemplate
    }
})
$ExportTemplate.Add_Click({
    $OpenFileDialog = New-Object Microsoft.Win32.SaveFileDialog
    $OpenFileDialog.InitialDirectory = $LastLocation
    $OpenFileDialog.AddExtension = $true
    $OpenFileDialog.DefaultExt = "json"
    $OpenFileDialog.Filter = "Json Template (*.json)|*.json"
    if($OpenFileDialog.ShowDialog()){
        $SaveLocation = $OpenFileDialog.FileName
        Write-Host "$SaveLocation" -ForegroundColor Yellow
        if($SaveLocation){
            Export-Template -path $SaveLocation
        }
    }
})
function DropdownRefineAction {
    $Already = $this.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}
    #$list = $this.parent.parent.children[0].children[1].tag | ?{$Already -notcontains $_}

    $ListFull = $this.parent.parent.tag | ?{$Already -notcontains $_}
    $ListFull | %{New-Object System.Windows.Controls.ComboBoxItem -Property @{Content=$_}} | %{$this.Items.Add($_)} | Out-Null
}

$DistributionPointCB.StaysOpenOnEdit = $true
$DistributionPointCB.add_Keyup({
    #$Already = $this.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}
    #$list = $this.tag | ?{$Already -notcontains $_}
    SearchBoxActionKeyup #-FullList $this.tag
})
$DistributionPointCB.add_Keydown({
    SearchBoxActionKeydown
})
$DistributionPointGroupCB.add_Keyup({
    #$Already = $this.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}
    #$list = $this.tag | ?{$Already -notcontains $_}
    SearchBoxActionKeyup #-FullList $this.tag
})
$DistributionPointGroupCB.StaysOpenOnEdit = $true
$DistributionPointGroupCB.add_Keydown({
    SearchBoxActionKeydown
})
$DetectionTypeCB.add_SelectionChanged({
    #Write-Host $DetectionTypeCB.SelectedItem.Content -ForegroundColor Green
    switch ($DetectionTypeCB.SelectedItem.Content) {
        "WindowsInstaller" {
            $DetectionMethodWISP.Visibility = 0
            $DetectionMethodFSSP.Visibility = 2
            $DetectionMethodREGSP.Visibility = 2
        }
        "FileSystem" {
            $DetectionMethodWISP.Visibility = 2
            $DetectionMethodFSSP.Visibility = 0
            $DetectionMethodREGSP.Visibility = 2
        }
        "Registry" {
            $DetectionMethodWISP.Visibility = 2
            $DetectionMethodFSSP.Visibility = 2
            $DetectionMethodREGSP.Visibility = 0
        }
        "App-V / will do later"{
            $DetectionMethodWISP.Visibility = 2
            $DetectionMethodFSSP.Visibility = 2
            $DetectionMethodREGSP.Visibility = 2
        }
    }
})
$AddToSCCM.Add_Click({
    $Splat = @{
        ApplicationDOIT=$ApplicationCH.IsChecked
        DeploymentTypeDOIT=$DeploymentTypeCH.IsChecked
        SuperDependDOIT=$SupDependencyCH.IsChecked
        DistributionDOIT=$DistributionCH.IsChecked
        CollectionDOIT=$CollectionCH.IsChecked
        DeploymentDOIT=$DeploymentCH.IsChecked
        ADGroupDOIT=$ADGroupsCH.IsChecked
    }
    ImportSCCM @Splat
})
$AddApplication.Add_Click({
    if(Check-App){ImportSCCM -ApplicationDOIT}
})
$AddDeploymentType.Add_Click({
    if(Check-DeploymentType){ImportSCCM -DeploymentTypeDOIT}
})
$AddDistribution.Add_Click({
    if(Check-Distribution){ImportSCCM -DistributionDOIT}
})
$CollectionDeploymentBT.Add_Click({
    if($ADGroupsCH.IsChecked){ImportSCCM -CollectionDOIT -ADGroupDOIT}
    else{ImportSCCM -CollectionDOIT}
})
$DeploymentBT.Add_Click({
    ImportSCCM -DeploymentDOIT
})
$SupersedenceButton.Add_Click({
    ImportSCCM -SuperDependDOIT
})
$ClearGUI.Add_Click({
    Start-Up -Clear
})
$AddDetectionMethodBT.Add_Click({
    Add-DetectionMethod
    $DeploymentTypeScrollViewer.PageDown()
})
$CreateImportBundle.Add_Click({
    Write-Host $PSCommandPath
    $Template = Export-Template -Pass
    $TemplateCheck = $true
    if($TemplateCheck){
        $OpenFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $OpenFileDialog.InitialDirectory = $LastLocation
        $OpenFileDialog.FileName = $Template.name
        if($OpenFileDialog.ShowDialog()){
            $SaveLocation = $OpenFileDialog.FileName
            Write-Host "$SaveLocation" -ForegroundColor Yellow
            if($SaveLocation){
                New-Item -Path $SaveLocation -ItemType directory
                Copy-Item -Path $PSCommandPath -Destination "$SaveLocation\ImportApp.ps1"
                
                $Command = "powershell.exe -executionpolicy bypass -f `"%~dp0ImportApp.ps1`" -TemplatePath `"$($Template.name).json`""
                $Zipped = Zip-Unzip -ZipPath "$SaveLocation\Template $($Template.name).zip" -FolderPath $Template.DeploymentType.CopyFrom
                
                if($Zipped){
                    $Template.DeploymentType.CopyFrom = ".\$($Zipped -replace '^.*\\','')"
                }
                $Command += " -Silent"
                #if($ApplicationCH.IsChecked)   {$Command += " -Application"}
                #if($DeploymentTypeCH.IsChecked){$Command += " -DeploymentType"}
                #if($DistributionCH.IsChecked)  {$Command += " -Distribution"}
                #if($CollectionCH.IsChecked)    {$Command += " -Collection"}
                #if($DeploymentCH.IsChecked)    {$Command += " -Deployment"}
                #if($ADGroupsCH.IsChecked)      {$Command += " -ADgroups"}
                $Template |  ConvertTo-Json -Depth 99 | New-Item -Path "$SaveLocation\Template $($Template.name).json" -ItemType File -force 
                New-Item -Path "$SaveLocation\ImportApp.cmd" -Value $Command
                if(Test-Path $Template.DeploymentType.CopyFrom){
                    #$Content = get
                }

            }
        }
    }
})

#endregion
#region Startup
$Global:LastLocation = $PSScriptRoot

$TabC.Height = 652
$TabC.HorizontalAlignment = 'Stretch'
$TabC.VerticalAlignment = 'Stretch'

if($ConnectedRCM -or (Connect-RCM)){
    [array]$FullList = Get-RCMApp -Name * -Fast | %{$_.LocalizedDisplayName} | Sort-Object
    Set-RComboBoxOptions -ItemList $FullList -ComboBox $SCCMSearchBoxCB -ReturnAction $IMPORT_SCCM_ACTION
    Set-RComboBoxOptions -ItemList $FullList -ComboBox $SupersedenceCB -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}
    Set-RComboBoxOptions -ItemList $FullList -ComboBox $DependencyCB -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}

    $DependencyCB.parent.parent.tag = $FullList
    $SupersedenceCB.parent.parent.tag = $FullList

    [array]$DistributionPointGroups = get-RCMDistributionPointGroup | %{$_.name}
    Set-RComboBoxOptions -ItemList $DistributionPointGroups -ComboBox $DistributionPointGroupCB  -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}

    [array]$DistributionPoints = get-RCMDistributionPoint | %{if($_.ItemName -match '\[\"Display=[\\]*([^\"]*?)[\\]*\"\]'){$Matches[1]}}
    Set-RComboBoxOptions -ItemList $DistributionPoints -ComboBox $DistributionPointCB  -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}
} 
else{
    [array]$FullList = 1..1000 | %{Get-Random} | Sort-Object

    Set-RComboBoxOptions -ItemList $FullList -ComboBox $SCCMSearchBoxCB -ReturnAction $IMPORT_SCCM_ACTION
    Set-RComboBoxOptions -ItemList $FullList -ComboBox $SupersedenceCB -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}
    Set-RComboBoxOptions -ItemList $FullList -ComboBox $DependencyCB -FilterLogic {$thisCB.Parent.Parent.Children | %{$_.Children[1].text} | ?{$_ -match "\S"}}
    
    $SupersedenceCB.parent.parent.tag = $FullList
    $DependencyCB.parent.parent.tag = $FullList

    [array]$DistributionPointGroups = get-RCMDistributionPointGroup | %{$_.name}

    $DistributionPointGroups | %{New-Object System.Windows.Controls.ComboBoxItem -Property @{Content=$_}} | %{$DistributionPointGroupCB.Items.Add($_)} | %{Out-Null}
    $DistributionPointGroupCB.tag = $DistributionPointGroups
    [array]$DistributionPoints = get-RCMDistributionPoint | %{if($_.ItemName -match '\[\"Display=[\\]*([^\"]*?)[\\]*\"\]'){$Matches[1]}}
    $DistributionPoints | %{New-Object System.Windows.Controls.ComboBoxItem -Property @{Content=$_}} | %{$DistributionPointCB.Items.Add($_)} | %{Out-Null}

    $DistributionPointCB.tag = $DistributionPoints
}

Start-Up
Check-App
Check-DeploymentType

$async = $window.Dispatcher.InvokeAsync({$window.ShowDialog() | Out-Null})
$global:RunningGUI = $true
$async.Wait() | Out-Null
$global:RunningGUI = $false


#endregion
}

#Connect Before anything else
Connect-RCM -Silent
if($Silent -and $TemplatePath){
    if(Test-Path $TemplatePath){
        $Template = Get-Content -Path $TemplatePath -Raw | ConvertTo-Json
    }
    elseif(Test-Path "$PSScriptRoot\$TemplatePath"){
        $Template = Get-Content -Path "$PSScriptRoot\$TemplatePath" -Raw | ConvertTo-Json
    }
    else{
        Write-Error -Message "template not found"
        Start-Sleep -Seconds 10
        exit 45
    }

    $Splat = @{
        ApplicationDOIT = $Template.AddApplication
        DeploymentTypeDOIT = $Template.AddDeploymentType
        DistributionDOIT = $Template.DistributeSource
        CollectionDOIT = $Template.AddCollection
        DeploymentDOIT = $Template.AddDeployment
        ADGroupDOIT = $Template.AddAdGroup
    }
    if(Connect-RCM -Silent){
        ImportSCCM @Splat
    }
    else{
        Write-Host "Cannot Connect to SCCM" -ForegroundColor Red
        Start-Sleep -Seconds 10
    }
    <#
    if(Connect-RCM -Silent){
        if($Application){
            $ApplicationObject = New-RCMAppFromTemplate -template $Template
        }
        else{
            $ApplicationObject = Get-RCMApp -Name $Template.Name
        }
        if($DeploymentType -and $ApplicationObject){
            $DeploymentTypeObject = New-RCMDeploymentTypeFromTemplate -Template $Template -Application $ApplicationObject
        }
        else{
            $DeploymentTypeObject = Get-RCMDeploymentType -ApplicationName $Template.Name -DeploymentTypeName *
        }
        if($ApplicationObject -and $DeploymentTypeObject -and $Distribution){
            $DistrubutionSplat = $Template.Distrubution | ConvertTo-Hashtable -RemoveNull
            $DistributionObject = Add-RCMDistribution -APPName $ApplicationObject.LocalizedDisplayName @DistrubutionSplat
        }
        if($Collection){
            if($ADgroups){$collectionObject = New-RCMCollectionFromTemplate -Template $Template -Application $ApplicationObject -Group}
            else{$collectionObject = New-RCMCollectionFromTemplate -Template $Template -Application $ApplicationObject}
            #foreach(){}
        }
        else{
            $CollectionName = $Template.Collection | %{$_.name} 
            if(!$CollectionName){$CollectionName = $Template.Name}
            $collection = $CollectionName | %{Get-RCMCollection -Name $_}
        }
        if($Deployment -and $ApplicationObject -and $DeploymentTypeObject -and $collectionObject){
            New-RCMDeploymentFromTemplate -ApplicationName $ApplicationObject.LocalizedDisplayName -Template $Template
        }
        Start-Sleep -Seconds 10
    }\Desktop\_SCCM\Application Deployment\2 - Business\Microsoft\Power BI Desktop\2.76\R01
    #>

}
else{
    if(!$TemplatePath){
        Call-RCMGui
    }
    elseif(Test-Path $TemplatePath){
        Call-RCMGui -TemplatePath $TemplatePath
    }
    elseif(Test-Path "$PSScriptRoot\$TemplatePath"){
        Call-RCMGui -TemplatePath "$PSScriptRoot\$TemplatePath"
    }
}
