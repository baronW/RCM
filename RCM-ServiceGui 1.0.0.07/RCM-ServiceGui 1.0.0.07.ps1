function Call-RCMServiceGui {
    $Blob = 300
    [xml]$ButtonXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Rory's Sccm Tools" WindowStartupLocation = "CenterScreen"
    SizeToContent = "WidthAndHeight" 
    ShowInTaskbar = "True" 
    Background = "White" 
    ResizeMode = "NoResize"
>
    <StackPanel Orientation = "Vertical" FlowDirection = "LeftToRight" Name="MainStackPanel">
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" FontWeight="Bold" Content = "App Details (wild card is *)" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Content = "App Name" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="GetAppDetailsCB" Height = "26" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "GetAppDetailsBT" Height = "26" Content = "Get App Details" ToolTip = "Get App Details" Width = "$Blob"/>
        </StackPanel>

        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26"  FontWeight="Bold" Content = "Enforcement State" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Content = "App Name" Width = "$Blob"/>
            <Label Height = "26" Content = "Collection Name" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="GetEnforcementStateCB1" Height = "26" Width = "$Blob"/>
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="GetEnforcementStateCB2" Height = "26" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "GetEnforcementStateBT" Height = "26" Content = "Get EnforcementState" ToolTip = "Get EnforcementState" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" FontWeight="Bold" Content = "Enforcement State Device Centric" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "GenerateNamesBUT2" Height = "26" Content = "Device Name (click to GenerateNames)" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="GetEnforcementStateDeviceCB1" Height = "26" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "GetDeviceEnforcementState" Height = "26" Content = "Get Device Enforcement State" ToolTip = "Get Device Enforcement State" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" FontWeight="Bold" Content = "Invoke Cycle" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "GenerateNamesBUT" Height = "26" Content = "Device Name (click to GenerateNames)" Width = "$Blob"/>
            <Label Height = "26" Content = "Cycle Name" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="CycleDevice" Height = "26" Width = "$Blob"/>
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="CycleName" Height = "26" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "CycleBT" Height = "26" Content = "Run Cycle" ToolTip = "Run Cycle" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Label Height = "26" Content = "Device Colection Name" Width = "$Blob"/>
            <Label Height = "26" Content = "Cycle Name" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="CycleColection" Height = "26" Width = "$Blob"/>
            <ComboBox IsTextSearchEnabled="false" IsEditable="true" Name="CycleName2" Height = "26" Width = "$Blob"/>
        </StackPanel>
        <StackPanel Orientation = "Horizontal" FlowDirection = "LeftToRight" >
            <Button x:Name = "CycleCBT2" Height = "26" Content = "Run Cycle" ToolTip = "Run Cycle" Width = "$Blob"/>
        </StackPanel>
    </StackPanel>
</Window>
"@
    #if($RCMisConnected -or $(Connect-RCM)){
        Add-Type -AssemblyName PresentationFramework
        $ButtonReader=(New-Object System.Xml.XmlNodeReader $ButtonXaml)
        $ButtonWindow=[Windows.Markup.XamlReader]::Load($ButtonReader)

        $MainStackPanel  = $ButtonWindow.FindName('MainStackPanel')
        $GetAppDetailsBT = $ButtonWindow.FindName('GetAppDetailsBT')
        $GetEnforcementStateBT = $ButtonWindow.FindName('GetEnforcementStateBT')
        $GetAppDetailsCB = $ButtonWindow.FindName('GetAppDetailsCB')
        $GetEnforcementStateCB1 = $ButtonWindow.FindName('GetEnforcementStateCB1')
        $GetEnforcementStateCB2 = $ButtonWindow.FindName('GetEnforcementStateCB2')
        $CycleDevice = $ButtonWindow.FindName('CycleDevice')
        $CycleName = $ButtonWindow.FindName('CycleName')
        $CycleBT = $ButtonWindow.FindName('CycleBT')
        $GenerateNamesBUT = $ButtonWindow.FindName('GenerateNamesBUT')
        $GetEnforcementStateDeviceCB1 = $ButtonWindow.FindName('GetEnforcementStateDeviceCB1')
        $GenerateNamesBUT2 = $ButtonWindow.FindName('GenerateNamesBUT2')
        $GetDeviceEnforcementState = $ButtonWindow.FindName('GetDeviceEnforcementState')
        $CycleColection = $ButtonWindow.FindName('CycleColection')
        $CycleName2 = $ButtonWindow.FindName('CycleName2')
        $CycleCBT2 = $ButtonWindow.FindName('CycleCBT2')

        $GetAppInfoAction = {
            if($GetAppDetailsCB.Text){
                $AppInfo = Get-RCMAppInfo -ApplicationName $GetAppDetailsCB.Text
            }
            else{
                $AppInfo = Get-RCMAppInfo -ApplicationName *
            }
            try{Export-RExcel -ErrorAction Stop -InputObject $AppInfo}
            catch{$AppInfo | Out-GridView}
        }

        $GetAppDetailsBT.add_Click({
            &$GetAppInfoAction
        })

        $GetEnforcementStateAction = {
            if($GetEnforcementStateCB1.Text -or $GetEnforcementStateCB2.Text){
                $EnforcementInfo = Get-RcmEnforcementState -ApplicationName $GetEnforcementStateCB1.Text -CollectionName $GetEnforcementStateCB2.Text
            }
            try{Export-RExcel -ErrorAction Stop -InputObject $EnforcementInfo}
            catch{$EnforcementInfo | Out-GridView}
        }
        $GetDeviceEnforcementStateAction = {
            if($GetEnforcementStateDeviceCB1.Text){
                $EnforcementInfoD = Get-RcmEnforcementStateDevice -DeviceName $GetEnforcementStateDeviceCB1.Text
            }
            try{Export-RExcel -ErrorAction Stop -InputObject $EnforcementInfoD}
            catch{$EnforcementInfoD | Out-GridView}
        }

        $InvokeCycleAction = {
            Invoke-RCMCycle -ComputerName $CycleDevice.Text -Cycle $CycleName.text
        }
        $InvokeCycleAction2 = {
            Write-Verbose $CycleColection.text
            $Members = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "SELECT * FROM SMS_Collection WHERE Name = '$($CycleColection.text)'" | %{Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "SELECT * FROM $($_.MemberClassName)"} | %{$_.Name} | ?{$_}
            foreach($D in $Members){
                Invoke-RCMCycle -ComputerName $D -Cycle $CycleName.text
            }
        }
        $CycleBT.add_Click({
            &$InvokeCycleAction
        })
        $CycleCBT2.add_Click({
            &$InvokeCycleAction2
        })

        $GetEnforcementStateBT.add_Click({
            &$GetEnforcementStateAction
        })
        
        $GetDeviceEnforcementState.add_Click({
            &$GetDeviceEnforcementStateAction
        })

        [array]$FullList = Get-RCMApp -Name * -Fast | %{$_.LocalizedDisplayName} | Sort-Object
        [array]$ALLCollectionsO = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "SELECT * FROM SMS_Collection"
        [array]$Collections = $ALLCollectionsO | %{$_.name}
        [Array]$DeviceCollections = $ALLCollectionsO | ?{$_.CollectionType -eq "2"} | %{$_.name}

        Set-RComboBoxOptions -ComboBox $GetAppDetailsCB -ItemList $FullList -ReturnAction $GetAppInfoAction -OpenOnSelect
        Set-RComboBoxOptions -ComboBox $GetEnforcementStateCB1 -ItemList $FullList -ReturnAction $GetEnforcementStateAction -OpenOnSelect
        Set-RComboBoxOptions -ComboBox $GetEnforcementStateCB2 -ItemList $Collections -ReturnAction $GetEnforcementStateAction -OpenOnSelect
        Set-RComboBoxOptions -ComboBox $CycleColection -ItemList $DeviceCollections -ReturnAction $InvokeCycleAction2 -OpenOnSelect


        $GENBUT = {
            $Devices = Get-WmiObject -Query "SELECT * FROM SMS_CM_RES_COLL_SMS00001" -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName | %{$_.name}
            Set-RComboBoxOptions -ComboBox $CycleDevice -ItemList $Devices -ReturnAction $GetDeviceEnforcementStateAction -OpenOnSelect
            Set-RComboBoxOptions -ComboBox $GetEnforcementStateDeviceCB1 -ItemList $Devices -ReturnAction $GetEnforcementStateAction -OpenOnSelect
        }
        $GenerateNamesBUT.add_Click({
            &$GENBUT
        })
        $GenerateNamesBUT2.add_Click({
            &$GENBUT
        })
        
        $ActionList = "Application Deployment Evaluation Cycle","Discovery Data Collection Cycle","Hardware Inventory Cycle","Machine Policy Retrieval and Evaluation Cycle","Software Inventory Cycle","Software Metering Usage Report Cycle","Software Updates Deployment Evaluation Cycle","Software Updates Scan Cycle","Windows Installer Source List Update Cycle","Machine Retrieval & Application Deployment"
        Set-RComboBoxOptions -ComboBox $CycleName -ItemList $ActionList -ReturnAction $InvokeCycleAction -SelectedIndex 9 -OpenOnSelect
        Set-RComboBoxOptions -ComboBox $CycleName2 -ItemList $ActionList -ReturnAction $InvokeCycleAction2 -SelectedIndex 9 -OpenOnSelect

        $async = $ButtonWindow.Dispatcher.InvokeAsync({$ButtonWindow.ShowDialog()})
        $async.Wait()

        $Global:Output_3ae4595be1c84e0db1e08b30491023e9 
    #}
}
function Get-RCMAppInfo {
    param(
        [string]$ApplicationName,
        [switch]$OutGridView,
        $appID,
        $ModelName
    )
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $Global:RCMSiteCode
        $Splat = @{ComputerName="$Global:ProviderMachineName";Namespace="ROOT\SMS\Site_$Global:RCMSiteCodeRaw"}
        if ($ModelName){$Apps = Get-CMApplication -ModelName $ModelName}
        elseif ($appID){$Apps = Get-CMApplication -Id $appID}
        else{$Apps = Get-CMApplication -Name $ApplicationName}
        
        $ApplicationMatchingHash = @{}
        $Apps | %{$ApplicationMatchingHash += @{$_.ModelName = $_.LocalizedDisplayName}}

        $TrueOutput = $null

        #$QUERYNAME = $Apps.LocalizedDisplayName -join "','"
        #$ALLDeploymentInfo = Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName IN ('$QUERYNAME')"

        $TrueOutput = foreach ($App in $Apps){
            $AppXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($App.SDMPackageXML)

            switch ($AppXML.HighPriority) {
                '1' {$Distributionpriority = 'High'}
                '2' {$Distributionpriority = 'Medium'}
                '3' {$Distributionpriority = 'Low'}
                Default {$Distributionpriority = 'Medium'}
            }
            #$OutPut = New-Object -TypeName psobject

            $Supersedes = foreach ($deployment in $AppXml.DeploymentTypes){
                Foreach ($Supersede in $deployment.Supersedes){
                    foreach ($Opa in $Supersede.Expression){
                        $ModelName = "$($Opa.ApplicationAuthoringScopeId)/$($Opa.ApplicationLogicalName)"
                        if ($ApplicationMatchingHash.$ModelName -ne $null){
                            $ApplicationMatchingHash.$ModelName
                        }
                        else{
                            $OUTAppName = $(Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName -query "select LocalizedDisplayName from SMS_Application where ModelName='$ModelName'" ).LocalizedDisplayName
                            $ApplicationMatchingHash += @{$ModelName = $OUTAppName}
                        }
                    }
                }
            }

            #$DeploymentInfo = $ALLDeploymentInfo | ?{$_.TargetName -eq $App.LocalizedDisplayName} #Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName='$($App.LocalizedDisplayName)'"
            $DeploymentInfo = Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName='$($App.LocalizedDisplayName)'"

            foreach ($deployment in $AppXml.DeploymentTypes){
                $OB = New-Object -TypeName psobject
                $Tecnology = $deployment.Technology
                $OB | Add-Member -MemberType NoteProperty -Name 'name' -Value $App.LocalizedDisplayName
                $OB | Add-Member -MemberType NoteProperty -Name 'DeploymentTypeName' -Value $deployment.Title
                $OB | Add-Member -MemberType NoteProperty -Name 'ContentLocation' -Value $($deployment.Installer.Contents | %{$_.location})
                $OB | Add-Member -MemberType NoteProperty -Name 'FolderPath' -Value $(Get-RCMAppFolderPath -AppName $App.LocalizedDisplayName)

                $CollectionInformation = foreach($D in $DeploymentInfo){
                    $CollectionInfo = New-Object -TypeName psobject
                    $CollectionInfo | Add-Member -NotePropertyName "CollectionName" -NotePropertyValue $D.CollectionName
                    $Collection = Get-CMCollection -Id $D.CollectionID
                    switch ($Collection.CollectionType){
                        1 {
                            $CollectionType = "UserCollection"
                        }
                        2 {
                            $CollectionType = "DeviceCollection"
                        }
                    }
                    $CollectionInfo | Add-Member -NotePropertyName "CollectionType" -NotePropertyValue $CollectionType
                    $CollectionInfo | Add-Member -NotePropertyName "Purpose" -NotePropertyValue $D.TargetSubName
                    $CollectionInfo | Add-Member -NotePropertyName "RuleNames" -NotePropertyValue $($Collection | %{$_.CollectionRules | %{$_.RuleName}})
                    $CollectionInfo | Add-Member -NotePropertyName "QueryExpression" -NotePropertyValue $($Collection | %{$_.CollectionRules | %{$_.QueryExpression}})
                    $CollectionInfo
                }

                $UserInstall = $CollectionInformation | ?{($_.CollectionType -eq "UserCollection") -and ($_.Purpose -eq "Install")}
                $UserUninstall = $CollectionInformation | ?{($_.CollectionType -eq "UserCollection") -and ($_.Purpose -eq "Remove")}
                $DeviceInstall = $CollectionInformation | ?{($_.CollectionType -eq "DeviceCollection") -and ($_.Purpose -eq "Install")}
                $DeviceUninstall = $CollectionInformation | ?{($_.CollectionType -eq "DeviceCollection") -and ($_.Purpose -eq "Remove")}

                $OB | Add-Member -MemberType NoteProperty -Name 'Collections' -Value $($CollectionInformation | %{$_.CollectionName})
                $OB | Add-Member -MemberType NoteProperty -Name 'InstallRules' -Value $($CollectionInformation | ?{($_.Purpose -eq "Install")} | %{$_.RuleNames})
                $OB | Add-Member -MemberType NoteProperty -Name 'RemoveRules' -Value $($CollectionInformation | ?{($_.Purpose -eq "Remove")} | %{$_.RuleNames})

                
                $OB | Add-Member -MemberType NoteProperty -Name 'InstallCommandLine' -Value $deployment.Installer.InstallCommandLine
                $OB | Add-Member -MemberType NoteProperty -Name 'UninstallCommandLine' -Value $deployment.Installer.InstallCommandLine
                $OB | Add-Member -MemberType NoteProperty -Name 'RepairCommandLine' -Value $deployment.Installer.RepairCommandLine
                $OB | Add-Member -MemberType NoteProperty -Name 'Tecnology' -Value $Tecnology
                $OB | Add-Member -MemberType NoteProperty -Name 'IsEnabled' -Value $App.IsEnabled
                $OB | Add-Member -MemberType NoteProperty -Name 'DeploymentName' -Value $DeploymentInfo.DeploymentName
                $OB | Add-Member -MemberType NoteProperty -Name 'ModelName' -Value $App.ModelName
                $OB | Add-Member -MemberType NoteProperty -Name 'CI_UniqueID' -Value $App.CI_UniqueID
                $OB | Add-Member -MemberType NoteProperty -Name 'Is Enabled' -Value $App.IsEnabled
                $OB | Add-Member -MemberType NoteProperty -Name 'Devices With App' -Value $App.NumberOfDevicesWithApp
                $OB | Add-Member -MemberType NoteProperty -Name 'Publisher' -Value $App.Manufacturer
                $OB | Add-Member -MemberType NoteProperty -Name 'Software version' -Value $APP.SoftwareVersion
                $OB | Add-Member -MemberType NoteProperty -Name 'Allow this application to be install form task sequence' -Value $(if($AppXML.AutoInstall){$true}else{$false}) #I think this is wrong
                $OB | Add-Member -MemberType NoteProperty -Name 'Application catalog Localized application name' -Value $($AppXML.DisplayInfo | %{$_.title})
                $OB | Add-Member -MemberType NoteProperty -Name "Is Superseded" -Value $App.IsSuperseded
                $OB | Add-Member -MemberType NoteProperty -Name "Is Superseding" -Value $App.IsSuperseding
                $OB | Add-Member -MemberType NoteProperty -Name 'Superseds' -Value $Supersedes
                $OB | Add-Member -MemberType NoteProperty -Name 'Distribution priority' -Value $Distributionpriority
                $Dependencies = Foreach ($Dependency in $deployment.Dependencies){
                    foreach ($Opa in $Dependency.Expression.Operands){
                        #$(Get-WmiObject @Splat -query "select LocalizedDisplayName from SMS_Application where ModelName='$($Opa.ApplicationAuthoringScopeId)/$($Opa.ApplicationLogicalName)'" ).LocalizedDisplayName
                        $ModelName = "$($Opa.ApplicationAuthoringScopeId)/$($Opa.ApplicationLogicalName)"
                        if ($ApplicationMatchingHash.$ModelName -ne $null){
                            $ApplicationMatchingHash.$ModelName
                        }
                        else{
                            $OUTAppName = $(Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $Global:ProviderMachineName -query "select LocalizedDisplayName from SMS_Application where ModelName='$ModelName'" ).LocalizedDisplayName
                            $ApplicationMatchingHash += @{$ModelName = $OUTAppName}
                        }
                    }
                }
                $OB | Add-Member -Name 'Dependencies' -Value $Dependencies -MemberType NoteProperty
                $OB | Add-Member -Name 'Requirements' -Value $($deployment.Requirements | %{$_.name}) -MemberType NoteProperty
                $OB | Add-Member -Name 'PublishedShortcuts' -Value $($deployment.Installer.AvailableApplications | %{if ($_.Publish){$_.name}}) -MemberType NoteProperty
                $OB | Add-Member -Name 'PublishingInformation' -Value $deployment.Installer.AvailableApplications -MemberType NoteProperty
                #$OB | Add-Member -Name 'InstallCommandLine' -Value $deployment.Installer.InstallCommandLine -MemberType NoteProperty
                $OB | Add-Member -Name 'InstallFolder' -Value $deployment.Installer.InstallFolder -MemberType NoteProperty
                #$OB | Add-Member -Name 'UninstallCommandLine' -Value $deployment.Installer.UninstallCommandLine -MemberType NoteProperty
                $OB | Add-Member -Name 'UninstallFolder' -Value $deployment.Installer.UninstallFolder -MemberType NoteProperty
                $OB | Add-Member -Name 'ProductCode' -Value $deployment.Installer.ProductCode -MemberType NoteProperty
                $OB | Add-Member -Name 'InstallAs32Bit' -Value $deployment.Installer.RedirectCommandLine -MemberType NoteProperty
                $OB | Add-Member -Name 'ExecutionContext' -Value $deployment.Installer.ExecutionContext -MemberType NoteProperty
                $OB | Add-Member -Name 'RequiresLogOn' -Value $deployment.Installer.RequiresLogOn -MemberType NoteProperty
                $OB | Add-Member -Name 'InstallationProgramVisibility' -Value $deployment.Installer.UserInteractionMode -MemberType NoteProperty
                $OB | Add-Member -Name 'EstimatedRunTime' -Value $deployment.Installer.ExecuteTime -MemberType NoteProperty
                $OB | Add-Member -Name 'MaxRunTime' -Value $deployment.Installer.MaxExecuteTime -MemberType NoteProperty
                $OB | Add-Member -Name 'BehaviourSettings' -Value $deployment.Installer.PostInstallBehavior -MemberType NoteProperty
                $OB | Add-Member -Name 'DetectionMethodType' -Value $deployment.Installer.DetectionMethod -MemberType NoteProperty
            
                #This section Disintangels the Detection methid
                if ($deployment.Installer.DetectionMethod -eq 'Enhanced'){
                    [xml]$DeploymentXML =  [xml]$deployment.Installer.EnhancedDetectionMethod.Xml
                
                    $SettingNodes = $DeploymentXML.EnhancedDetectionMethod.Settings.ChildNodes
                
                    $RuleNode = $DeploymentXML.EnhancedDetectionMethod.Rule.Expression
                
                    $StringRule = @()
                    $Depth = 0
                    $count2 = 0
                    $RuleBlock = {
                        if (@('And','Or') -contains $CurrentRuleNode.Operator){
                            $StringRule += "$($CurrentRuleNode.Operator)"
                            $Depth++
                            Foreach ($node in $CurrentRuleNode.Operands.ChildNodes){
                                $CurrentRuleNode = $node
                                &$RuleBlock
                            }
                            $Depth--
                        }
                        else{
                            $RuleJoiners = $StringRule -join ' :: '
                            $RuleInnerText = $($SettingNodes[$count2].InnerText)
                            $RuleType = $(if ($SettingNodes[$count2].LogicalName -match '(.*)_'){$Matches[1]})

                            if ($CurrentRuleNode.Operator -ne 'NotEquals'){
                            
                                $ComplexRule = "$($CurrentRuleNode.Operands.ChildNodes[1].DataType) $($CurrentRuleNode.Operator) $($CurrentRuleNode.Operands.ChildNodes[1].Value)"
                            
                                "$RuleJoiners :: $RuleType :: $RuleInnerText :: $ComplexRule".TrimStart(" :: ")
                            }
                            else{
                                "$RuleJoiners :: $RuleType :: $RuleInnerText".TrimStart(" :: ")
                            }
                            $count2++
                        }
                    }
                    $CurrentRuleNode = $RuleNode
                
                    $settings = &$RuleBlock
                }
                else{
                    $SettingsSource = ''
                    $SettingsOperator = ''
                    $settings = ''
                }
                $OB | Add-Member -Name 'DetectionMethod' -Value $settings -MemberType NoteProperty

                $OB | Add-Member -MemberType NoteProperty -Name 'UI Collection' -Value $($UserInstall | %{$_.CollectionName})
                $OB | Add-Member -MemberType NoteProperty -Name 'UI RuleNames' -Value $($UserInstall | %{$_.RuleNames})
                $OB | Add-Member -MemberType NoteProperty -Name 'UI Query' -Value $($UserInstall | %{$_.QueryExpression})

                $OB | Add-Member -MemberType NoteProperty -Name 'UR Collection' -Value $($UserUninstall | %{$_.CollectionName})
                $OB | Add-Member -MemberType NoteProperty -Name 'UR RuleNames' -Value $($UserUninstall | %{$_.RuleNames})
                $OB | Add-Member -MemberType NoteProperty -Name 'UR Query' -Value $($UserUninstall | %{$_.QueryExpression})

                $OB | Add-Member -MemberType NoteProperty -Name 'DI Collection' -Value $($DeviceInstall | %{$_.CollectionName})
                $OB | Add-Member -MemberType NoteProperty -Name 'DI RuleNames' -Value $($DeviceInstall | %{$_.RuleNames})
                $OB | Add-Member -MemberType NoteProperty -Name 'DI Query' -Value $($DeviceInstall | %{$_.QueryExpression})

                $OB | Add-Member -MemberType NoteProperty -Name 'DR Collection' -Value $($DeviceUninstall | %{$_.CollectionName})
                $OB | Add-Member -MemberType NoteProperty -Name 'DR RuleNames' -Value $($DeviceUninstall | %{$_.RuleNames})
                $OB | Add-Member -MemberType NoteProperty -Name 'DR Query' -Value $($DeviceUninstall | %{$_.QueryExpression})


                $OB
            }
        }

        if($OutGridView){$TrueOutput | Out-GridView}
        $TrueOutput

        Set-Location $ReturnLocaltion.Path
    }
}
function Export-RExcel {
    param(
            [Parameter(ValueFromPipeline=$true)][array]$InputObject,
            [string]$Path,
            [string]$SheetName,
            [__ComObject]$ExcelObject,
            [int32]$LimitColumnWidth,
            [switch]$Silent,
            [switch]$CloseCOM,
            [switch]$PassThru,
            [int]$InteriorColorIndex=49,
            [int]$FontColorIndex=2,
            [array]$Graph
    )

    begin{
        try{
            if($ExcelObject){
                $Excel = $ExcelObject
            }
            else{
                $ExcelPrime = New-Object -Com "excel.Application"
            }
        }
        catch{
            Write-Error "Excel not installed"
            break
        }
        $InputObjectP = @()
    }

    Process {
        $InputObjectP += $InputObject
    }
    end{
        $properties = $InputObjectP[0].psobject.properties.name



        $AZ = 65..90|foreach-object{[char]$_}
        $AAZZ = foreach ($Lettter in $AZ){
            foreach ($OtherLettter in $AZ){
                $Lettter + $OtherLettter
            }
        }

        $AAAZZZ = foreach ($Lettter in $AZ){
            foreach ($OtherLettter in $AAZZ){
                $Lettter + $OtherLettter 
            }
        }
        

        $AZZ = $AZ + $AAZZ + $AAAZZZ
        #$AZZ.Count

        $LastC = $AZZ[($properties.Count - 1)]

        $2Darray = New-Object 'object[,]' ($InputObjectP.count + 1),$properties.Count
        
        foreach ($N in 0..($properties.Count - 1)){
            $2Darray[0,$N] = $properties[$N]
        }
        $row = 1
        $Q = 0
        $P = 0
        foreach ($Object in $InputObjectP) #this means for each IP in the data
        {
            $P++
            Foreach ($N in (0..($properties.count - 1))){
                $Q++
                $2Darray[$row,$N] = $Object.$($properties[$N]) -join "," # | ForEach-Object {if($_){$_}}) -join ", "    #GotaGoFast
            }
            $row ++
        }
        Write-Verbose "$Q $P"
        if (!($Excel.WorkSheets)){
            $ExcelPrime.visible = $false
            #else{$ExcelPrime.visible = $false}
            Write-Verbose "Creating new sheet"
            $Excel = $ExcelPrime.Workbooks.Add()
            $NewBook = $true
            $Sheet = $Excel.Worksheets.Item(1)
            #$Sheet = $Excel.WorkSheets.Item(1)
            if ($Excel.WorkSheets.Count -gt 1){
                foreach ($S in $Excel.WorkSheets.Count..2){
                    $Excel.WorkSheets.Item($S).Delete()
                }
            }
        }
        else{
            $MissingType = [System.Type]::Missing
            $Sheet = $Excel.WorkSheets.Add($MissingType,$Excel.WorkSheets.Item($Excel.WorkSheets.Count))
            $NewBook = $false
        }
        $Sheet.Range("A1","$($LastC)1").Interior.ColorIndex = $InteriorColorIndex
        $Sheet.Range("A1","$($LastC)1").Font.ColorIndex = $FontColorIndex
        $Sheet.Range("A1","$($LastC)1").Font.Bold = $True

        $Sheet.Range("A1","$($LastC)$row").Value2 = $2Darray
        $Sheet.Range("A1","$($LastC)1").EntireColumn.AutoFilter().null | Out-Null
        $Sheet.Range("A1","$($LastC)1").EntireColumn.AutoFit().null | Out-Null
        if ($LimitColumnWidth){
            foreach ($X in $AZZ){
                if ($Sheet.Range("$($X)1","$($X)1").ColumnWidth -gt $LimitColumnWidth){
                    $Sheet.Range("$($X)1","$($X)1").ColumnWidth = $LimitColumnWidth
                }
                #$X
                if ($X -eq $LastC){
                    break
                }
            }
        }

        if ($SheetName){
            if (1..$Excel.Worksheets.Count | ?{$Excel.Worksheets.Item($_).name -eq $SheetName}){
                Write-Host "$SheetName is already in use" -ForegroundColor Yellow
            }
            else{
                try{$Sheet.Name = $SheetName}
                catch{Write-Host "Can not rename sheet to '$SheetName' check length" -ForegroundColor Yellow}
            }
        }

        if($Graph){
            Add-Type -AssemblyName Microsoft.Office.Interop.Excel
            $RowCol = New-Object -TypeName Microsoft.Office.Core.XlRowCol
            $RowCol.value__ = 2
            Write-Verbose "making pretty graphs"
            $MissingType = [System.Type]::Missing
            $GraphSheet = $Excel.WorkSheets.Add($MissingType,$Excel.WorkSheets.Item($Excel.WorkSheets.Count))
            $NameColumn = 
        
            $LastUsedRow = 0
            $Spacer = 2
            $LastEnd=0
            foreach($G in $Graph){
                [array]$Values = $InputObjectP | %{$_.$G} | Sort-Object -Unique
                $New2Darray = New-Object 'object[,]' ($Values.count +1),3
                for ($i = 0;($i -lt $properties.Count) -and ($properties[$i] -ne $G); $i++){}
                $ColumnIndex = $AZZ[$i]
                #=COUNTIF('ESRI ArcGIS for Desktop'!$L:$L,$A2)
                $New2Darray[0,0] = $G
                $New2Darray[0,1] = "Count"
                for ($R = 1; $R -le $Values.count; $R++){ 
                    $ThisRowIndex =  $LastUsedRow + $R + $Spacer
                    $New2Darray[$R,0] = "$($Values[$R-1])"
                    $New2Darray[$R,1] = [string]"=COUNTIF('$($Sheet.Name)'!`$$($ColumnIndex)2:`$$($ColumnIndex)$($2Darray.GetLongLength(0)),`$A$($ThisRowIndex))"
                    $New2Darray[$R,2] = [string]"=`$B$($ThisRowIndex)/$($2Darray.GetLongLength(0)-1)"
                    #"=COUNTIF('$($Sheet.Name)'!`$$($ColumnIndex)2:`$$($ColumnIndex)$($Values.count + 1),`$A$($ThisRowIndex))"
                }
            
                $FirstRow = $LastUsedRow+$Spacer
                $LastRow = $FirstRow + $Values.count

                #$FirstCell = "A$($LastUsedRow+$Spacer)"
                #$LastCell = "B$($R + $FirstCell)"
                $GraphSheet.Range("A$FirstRow","C$LastRow").Value2 = $New2Darray
                $GraphSheet.Range("C$FirstRow","C$LastRow").NumberFormat = "0.0%"
                $GraphSheet.Range("A$FirstRow","C$FirstRow").Interior.ColorIndex = $InteriorColorIndex
                $GraphSheet.Range("A$FirstRow","C$FirstRow").Font.ColorIndex = $FontColorIndex
                $LastUsedRow = $LastRow
                $Start = $GraphSheet.Range("A1","A$($FirstRow -1)").Height
                if($Start -lt $LastEnd){$Start = $LastEnd}
                $ChartObject = $GraphSheet.ChartObjects().Add(200, $Start, 500, 450)
                $ChartObject.Chart.SetSourceData($GraphSheet.Range("A$FirstRow","B$LastRow"),$RowCol)
                $ChartObject.Chart.ChartType = 5
                $ChartObject.Chart.hastitle = $true
                $ChartObject.Chart.ChartTitle.text = $G
                $ChartObject.Chart.PlotArea.Format.Glow.Color.RGB = 100
                $LastEnd = $Start + 450

                #$DC++
                #$ChartObject.Chart.SetSourceData($Sheet.Range("A$($Upto)","B$($RowM + $Upto - 1)"),$RowCol)
                #$ChartObject.Chart.ChartType = 5
                #$ChartObject.Chart.hastitle = $true
                #$ChartObject.Chart.ChartTitle.text = "Types of $($LookingFor.DisplayName) - $S - $(Get-Date -Format "MMMM")"

            }
            $GraphSheet.Range("A1","C1").EntireColumn.AutoFit().null | Out-Null
            if ($SheetName){
                if (1..$Excel.Worksheets.Count | ?{$Excel.Worksheets.Item($_).name -eq "$SheetName Stats"}){
                    Write-Host "'$SheetName Stats' is already in use" -ForegroundColor Yellow
                }
                else{
                    try{$Sheet.Name = "$SheetName Stats"}
                    catch{Write-Host "Can not rename sheet to '$SheetName Stats' check length" -ForegroundColor Yellow}
                }
            }
        }

        if (!$Silent){$Excel.Parent.Visible = $true}

        if ($Path){$Sheet.SaveAs($Path) | Out-Null}
        #New-Variable -Name "Rory's totally awesome excel COM object" -Force -Scope global -Value $Excel | Out-Null
        if($PassThru){
            $Excel
        }
        elseif($CloseCOM -or $Silent){
            $Excel.Close()
        }
    }
}
function Connect-RCM {
    param($Site)
    if($Site){$Site = $Site.trim('\').trim(':')}
    if((!$RCMisConnected) -or ($Site -and ($RCMSiteCodeRaw -ne $Site))){
        Try{
            Import-Module "$ENV:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1" -Scope Global -ErrorAction SilentlyContinue # Import the ConfigurationManager.psd1 module
            [array]$drive0 = Get-PSDrive | ?{$_.Provider.name -eq 'CMSite'}
            if ($drive0.count -gt 1){
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
function Get-RcmEnforcementState {
    param(
        [Parameter(Mandatory=$false)][string]$ApplicationName="%",
        [Parameter(Mandatory=$false)][string]$CollectionName=$null
    )
    if($RCMisConnected -or $(Connect-RCM)){
        
        if(!$($ApplicationName).Trim()){$ApplicationName="%"}
        $DeploymentInfo = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName='$ApplicationName'"
        if($CollectionName){
            $CNW=$CollectionName}
        else{
            $CNW="%"
        }
    
        $AppDeploymentAssetDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "Select * from SMS_AppDeploymentAssetDetails where AppName like '$ApplicationName' AND CollectionName like '$CNW'"
        if(!$DeploymentInfo){
            $ApplicationNameDerived = $AppDeploymentAssetDetails | %{$_.AppName} | Sort-Object -Unique
            $DeploymentInfo = $ApplicationNameDerived | %{Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName -Query "Select * from SMS_DeploymentInfo where TargetName='$_'"}
        }
        $AppDeploymentAssetDetailsHASH = @{}
        $AppDeploymentAssetDetails | %{$AppDeploymentAssetDetailsHASH.$($_.MachineName)=$_}

        $ReturnLocaltion = Get-Location
        Set-Location $RCMSiteCode

        $CollectionDevices0 = foreach($D in $DeploymentInfo){
            $CollectionObject = Get-CMCollection -Id $D.CollectionID
            if ($CollectionObject.CollectionType -eq 1){
                $MachineNames = $AppDeploymentAssetDetails | ?{$_.CollectionID -eq $D.CollectionID} | %{$_.MachineName}
                $Split = Split-array -inArray $MachineNames -size 256
                $SearchString = $Split |%{"'" + $($_.SyncRoot -join "','") + "'"}
                $SearchString | %{Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "SELECT * FROM SMS_CM_RES_COLL_SMS00001 WHERE Name is in ($_)" | 
                    Add-Member -NotePropertyName "CollectionName" -NotePropertyValue $D.CollectionName -PassThru |
                    Add-Member -NotePropertyName "AapplicationName" -NotePropertyValue $D.TargetName -PassThru
                }
            }
            else{
                Get-CMCollectionMember -CollectionId $D.CollectionID | 
                Add-Member -NotePropertyName "CollectionName" -NotePropertyValue $D.CollectionName -PassThru |
                Add-Member -NotePropertyName "ApplicationName" -NotePropertyValue $D.TargetName -PassThru
            }
        }
        [array]$CollectionDevices = $CollectionDevices0 | Sort-Object -Property Name
        
        Set-Location $ReturnLocaltion

        $AppDeploymentErrorAssetDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "Select * from SMS_AppDeploymentErrorAssetDetails where AppName like '$ApplicationName' AND CollectionName like '$CollectionName'"
        $AppDeploymentErrorAssetDetailsHASH = @{}
        $AppDeploymentErrorAssetDetails | %{$AppDeploymentErrorAssetDetailsHASH.$($_.MachineName)=$_}
        #foreach($CD in $CollectionDevices){
        $N=0
        #$CollectionDevices | %{
        for ($N = 0; $N -lt $CollectionDevices.count; $N++){ 
            $global:CD = $CollectionDevices[$N]

            #Write-Host "$N of $($CollectionDevices.count)"
            $A = $AppDeploymentAssetDetailsHASH.$($CD.name)  #$AppDeploymentAssetDetails | ?{$_.MachineName -eq $CD.name}
            #Write-Host "C0"
            $OUT = New-Object -TypeName psobject
            $OUT | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $CD.Name
            #$OUT | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $A.MachineName
            $OUT | Add-Member -MemberType NoteProperty -Name "Active" -Value $CD.IsActive
            $OUT | Add-Member -MemberType NoteProperty -Name "Client" -Value $CD.IsClient
            #IsActive,IsClient
            $OUT | Add-Member -MemberType NoteProperty -Name "UserName" -Value $A.UserName
            $OUT | Add-Member -MemberType NoteProperty -Name "ApplicationName" -Value $(if($CD.ApplicationName){$CD.ApplicationName}else{$A.AppName})
            $OUT | Add-Member -MemberType NoteProperty -Name "CollectionName" -Value $CD.CollectionName #$A.CollectionName
            $PU = $($CD.PrimaryUser -split ',') | %{if($_ -match ".*\\(.*)"){$Matches[1]}else{$_}} 
            $PUE = if($PU){$PU | %{try{Get-ADUser $_ -Properties EmailAddress -ErrorAction SilentlyContinue}catch{}} | %{$_.EmailAddress}}else{}
            $OUT | Add-Member -MemberType NoteProperty -Name "PrimaryUser" -Value $PU
            $OUT | Add-Member -MemberType NoteProperty -Name "PrimaryUserEmail" -Value $PUE
	        $LU = $($CD.LastLogonUser -split ',') | %{if($_ -match ".*\\(.*)"){$Matches[1]}else{$_}} 
            $LUE = if($LU){$LU | %{try{Get-ADUser $_ -Properties EmailAddress -ErrorAction SilentlyContinue}catch{}} | %{$_.EmailAddress}}else{}
            $OUT | Add-Member -MemberType NoteProperty -Name "LastLogonUser" -Value $LU
            $OUT | Add-Member -MemberType NoteProperty -Name "LastLogonUserEmail" -Value $LUE
            #Write-Host "C1"
        
            switch ($A.ComplianceState) {
                0 {$CS = "Compliance State Unknown"}
                1 {$CS = "Compliant"}
                2 {$CS = "Non-Compliant"}
                4 {$CS = "Error"}
                #$null {"Device offline"}
                default {$CS = "Unknown"} #$A.ComplianceState}
            }
            $OUT | Add-Member -MemberType NoteProperty -Name "ComplianceState" -Value $CS
            switch ($A.EnforcementState) {
            1000 {$ES = "Success"}
            1001 {$ES = "Already Compliant"}
            1002 {$ES = "Simulate Success"}
            2000 {$ES = "In progress"}
            2001 {$ES = "Waiting for content"}
            2002 {$ES = "Installing"}
            2003 {$ES = "Restart to continue"}
            2004 {$ES = "Waiting for maintenance window"}
            2005 {$ES = "Waiting for schedule"}
            2006 {$ES = "Downloading dependent content"}
            2007 {$ES = "Installing dependent content"}
            2008 {$ES = "Restart to complete"}
            2009 {$ES = "Content downloaded"}
            2010 {$ES = "Waiting for update"}
            2011 {$ES = "Waiting for user session reconnect"}
            2012 {$ES = "Waiting for user logoff"}
            2013 {$ES = "Waiting for user logon"}
            2014 {$ES = "Waiting To Install"}
            2015 {$ES = "Waiting Retry"}
            2016 {$ES = "Waiting For Presentation Mode"}
            2017 {$ES = "Waiting For Orchestration"}
            2018 {$ES = "Waiting For Network"}
            2019 {$ES = "Pending App-V Virtual Environment Update"}
            2020 {$ES = "Updating App-V Virtual Environment"}
            3000 {$ES = "Requirements not met"}
            3001 {$ES = "Host Platform Not Applicable"}
            4000 {$ES = "Unknown"}
            5000 {$ES = "Deployment failed"}
            5001 {$ES = "Evaluation failed"}
            5002 {$ES = "Deployment failed"}
            5003 {$ES = "Failed to locate content"}
            5004 {$ES = "Dependency installation failed"}
            5005 {$ES = "Failed to download dependent content"}
            5006 {$ES = "Conflicts with another application deployment"}
            5007 {$ES = "Waiting Retry"}
            5008 {$ES = "Failed to uninstall superseded deployment type"}
            5009 {$ES = "Failed to download superseded deployment type"}
            5010 {$ES = "Failed to updating App-V Virtual Environment"}
            default {$ES = $A.EnforcementState}
        }
            #Write-Host "C2"
            $OUT | Add-Member -MemberType NoteProperty -Name "EnforcementState" -Value $ES
            $OUT | Add-Member -MemberType NoteProperty -Name "ErrorCode" -Value ""
            if($AppDeploymentErrorAssetDetailsHASH.$($CD.name) -and $AppDeploymentErrorAssetDetailsHASH.$($CD.name).ErrorCode){
                $ErrorCode = "0x$([String]::Format("{0:x8}", $AppDeploymentErrorAssetDetailsHASH.$($CD.name).ErrorCode))"
                switch ($ErrorCode){
                    "0x87D00213" {$ErrorCode = "$ErrorCode Timeout occured"}
                    "0x87D00324" {$ErrorCode = "$ErrorCode The software package was not detected after instillation"}
                    "0x87D00325" {$ErrorCode = "$ErrorCode The software package uninstalled successfully, but a software detection rule was still found"}
                    "0x87D00607" {$ErrorCode = "$ErrorCode Unable to Download Software"}
                    "0x87d01201" {$ErrorCode = "$ErrorCode Unable to download install because hard drive is full"}
                    "0x80091007" {$ErrorCode = "$ErrorCode Hash value incorrect"}
                    "0x80041001" {$ErrorCode = "$ErrorCode SMS Agent Host service not running (ccmexec.exe)"}
                    "0x87D01106" {$ErrorCode = "$ErrorCode Content is not available on Distribution Point"}
                    "0x87D00443" {$ErrorCode = "$ErrorCode Cannot install because a conflicting process is running"}
                    "0x80040154" {$ErrorCode = "$ErrorCode Class not registered"}
                    "0x80041006" {$ErrorCode = "$ErrorCode Out of memory"}
                    "0x800706ba" {$ErrorCode = "$ErrorCode The RPC server is unavalible"}
                    "0x800706be" {$ErrorCode = "$ErrorCode The remote procedure call failed."}
                    "0x80070057" {$ErrorCode = "$ErrorCode The parameter is incorrect."}
                    "0x87d00231" {$ErrorCode = "$ErrorCode Transient error"}
                    "0x80041033" {$ErrorCode = "$ErrorCode Shutting down"}
                    "0x87d00317" {$ErrorCode = "$ErrorCode Unknown Error -2016410857"}
                    "0x800705b4" {$ErrorCode = "$ErrorCode The Timeout period expired"}
                    "0x87d0027c" {$ErrorCode = "$ErrorCode CI documents download timed out"}
                    "0x00000652" {$ErrorCode = "$ErrorCode Another installation is already in progress. Complete that installation before proceeding with this install"}
                    "0x80041009" {$ErrorCode = "$ErrorCode Not Available"}
                    default {$ES = $ErrorCode}
                }
                
                $OUT.ErrorCode = $ErrorCode
            }
            $OUT | Add-Member -MemberType NoteProperty -Name "DeviceOS" -Value $CD.DeviceOS
            #$OUT | Add-Member -MemberType NoteProperty -Name "LastLogonUser" -Value $CD.LastLogonUser
            $State = if($OUT.ErrorCode){$OUT.ErrorCode}elseif($ES){$ES}elseif($CS){$CS}elseif(!$CD.IsClient){"NoClient"}else{"Unknown"}
            $OUT | Add-Member -MemberType NoteProperty -Name "State" -Value $State
            $OUT | Select-Object -Property MachineName,ApplicationName,CollectionName,State,DeviceOS,ComplianceState,EnforcementState,ErrorCode,UserName,PrimaryUser,PrimaryUserEmail,LastLogonUser,LastLogonUserEmail,Active,Client	
        #$CD.IsClient.GetType()
        }
    }
}
function Invoke-RCMCycle {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][ValidateSet(
        "Application Deployment Evaluation Cycle",
        "Discovery Data Collection Cycle",
        "Hardware Inventory Cycle",
        "Machine Policy Retrieval and Evaluation Cycle",
        "Software Inventory Cycle",
        "Software Metering Usage Report Cycle",
        "Software Updates Deployment Evaluation Cycle",
        "Software Updates Scan Cycle",
        "Windows Installer Source List Update Cycle",
        "Machine Retrieval & Application Deployment"
        )][array]$Cycle
    )
    foreach($C in $Cycle){
        switch ($C)
        {
            "Application Deployment Evaluation Cycle"       {$CycleCode="{00000000-0000-0000-0000-000000000121}"}
            "Discovery Data Collection Cycle"               {$CycleCode="{00000000-0000-0000-0000-000000000003}"}
            "Hardware Inventory Cycle"                      {$CycleCode="{00000000-0000-0000-0000-000000000001}"}
            "Machine Policy Retrieval and Evaluation Cycle" {$CycleCode="{00000000-0000-0000-0000-000000000021}"}
            "Software Inventory Cycle"                      {$CycleCode="{00000000-0000-0000-0000-000000000002}"}
            "Software Metering Usage Report Cycle"          {$CycleCode="{00000000-0000-0000-0000-000000000031}"}
            "Software Updates Deployment Evaluation Cycle"  {$CycleCode="{00000000-0000-0000-0000-000000000108}"}
            "Software Updates Scan Cycle"                   {$CycleCode="{00000000-0000-0000-0000-000000000113}"}
            "Windows Installer Source List Update Cycle"    {$CycleCode="{00000000-0000-0000-0000-000000000032}"}
            "Machine Retrieval & Application Deployment"    {$CycleCode=@("{00000000-0000-0000-0000-000000000021}",@("{00000000-0000-0000-0000-000000000121}"))}
        }
        try{
            foreach($CC in $CycleCode){
                Write-Host "Running Cycle: $ComputerName : $C : $CC" -ForegroundColor Yellow
                Invoke-WMIMethod -ComputerName $ComputerName -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule -ArgumentList $CC -ErrorAction Stop | Out-Null
            }
        }
        catch{
            Write-Error $_
        }
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
        Set-Location $Global:RCMSiteCode
        $SPLAT = @{}
        if($Name){$SPLAT+=@{Name=$Name}}
        elseif($appID){$SPLAT+=@{appID=$appID}}
        elseif($ModelName){$SPLAT+=@{ModelName=$ModelName}}
        if($Fast){$SPLAT+=@{Fast=$true}}

        Get-CMApplication @SPLAT

        Set-Location $ReturnLocaltion.Path
    }
}
Function Set-RComboBoxOptions {
    Param([array]$ItemList,$ComboBox,$ReturnAction=$false,$FilterLogic,[switch]$OpenOnSelect,$SelectedIndex)
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
    if($OpenOnSelect){$ComboBox.add_GotFocus({if(!$this.IsDropDownOpen){$this.IsDropDownOpen = $true}})}
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
    if($SelectedIndex){$ComboBox.SelectedIndex = $SelectedIndex}
}
function Get-RCMAppFolderPath {
    param([Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][string]$AppName)
    if($RCMisConnected -or $(Connect-RCM)){
        $ReturnLocaltion = Get-Location
        Set-Location $Global:RCMSiteCode

        $SiteDetails = Get-CMSite
        $SiteServer = $Global:ProviderMachineName #$SiteDetails.ServerName
        $SiteCode =  $Global:RCMSiteCodeRaw #$SiteDetails.SiteCode
        $AppsToLookFor = Get-CMApplication -Name $AppName -fast

        if ($AppsToLookFor){
        Foreach ($app in $AppsToLookFor.LocalizedDisplayName){
                $Folder = Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $SiteServer -Query "Select * from SMS_ObjectContainerItem where ObjectType='6000' AND InstanceKey is in (Select ModelName from SMS_Application Where LocalizedDisplayName='$app')"
                $FolderDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $SiteServer -Query "select * from SMS_ObjectContainerNode where ObjectType=6000 And ContainerNodeID='$($Folder.ContainerNodeID)'"
                $FolderPath = "$($FolderDetails.Name)"
                while($FolderDetails.ParentContainerNodeID){
                    $FolderDetails = $FolderDetails=Get-WmiObject -Namespace "ROOT\SMS\Site_$Global:RCMSiteCodeRaw" -ComputerName $SiteServer -Query "select * from SMS_ObjectContainerNode where ObjectType=6000 And ContainerNodeID='$($FolderDetails.ParentContainerNodeID)'"
                    $FolderPath = "$($FolderDetails.Name)\$FolderPath"
                }
                "$SiteCode`:\Application\$FolderPath\$app"
            }
        }
        else{
            Write-Error -Message "Can't find an app named $AppName"
        }
        Set-Location $ReturnLocaltion.Path
    }
}
function Split-array {
  param($inArray,[int]$parts,[int]$size)
  
  if ($parts) {
    $PartSize = [Math]::Ceiling($inArray.count / $parts)
  } 
  if ($size) {
    $PartSize = $size
    $parts = [Math]::Ceiling($inArray.count / $size)
  }

  $outArray = @()
  for ($i=1; $i -le $parts; $i++) {
    $start = (($i-1)*$PartSize)
    $end = (($i)*$PartSize) - 1
    if ($end -ge $inArray.count) {$end = $inArray.count}
    $outArray+=,@($inArray[$start..$end])
  }
  return ,$outArray

}
function Get-RcmEnforcementStateDevice {
    param([string]$DeviceName)
    if($RCMisConnected -or $(Connect-RCM)){}else{return}
    $APPS = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "Select * from SMS_AppDeploymentAssetDetails where MachineName like '$DeviceName'" | Select-Object -Property AppName,CollectionName | Sort-Object -Property AppName
    foreach($A in $APPS){
        $AppDeploymentErrorAssetDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "Select * from SMS_AppDeploymentErrorAssetDetails where AppName like '$($A.AppName)' AND CollectionName like '$($A.CollectionName)' AND MachineName like '$DeviceName'"
        $AppDeploymentAssetDetails = Get-WmiObject -Namespace "ROOT\SMS\Site_$RCMSiteCodeRaw" -ComputerName $ProviderMachineName  -Query "Select * from SMS_AppDeploymentAssetDetails where AppName like '$($A.AppName)' AND CollectionName like '$($A.CollectionName)' AND MachineName like '$DeviceName'"
        switch ($AppDeploymentAssetDetails.ComplianceState) {
            0 {$CS = "Compliance State Unknown"}
            1 {$CS = "Compliant"}
            2 {$CS = "Non-Compliant"}
            4 {$CS = "Error"}
            #$null {"Device offline"}
            default {$CS = "Unknown"} #$A.ComplianceState}
        }
        $A | Add-Member -MemberType NoteProperty -Name "ComplianceState" -Value $CS -Force
        switch ($AppDeploymentAssetDetails.EnforcementState) {
            1000 {$ES = "Success"}
            1001 {$ES = "Already Compliant"}
            1002 {$ES = "Simulate Success"}
            2000 {$ES = "In progress"}
            2001 {$ES = "Waiting for content"}
            2002 {$ES = "Installing"}
            2003 {$ES = "Restart to continue"}
            2004 {$ES = "Waiting for maintenance window"}
            2005 {$ES = "Waiting for schedule"}
            2006 {$ES = "Downloading dependent content"}
            2007 {$ES = "Installing dependent content"}
            2008 {$ES = "Restart to complete"}
            2009 {$ES = "Content downloaded"}
            2010 {$ES = "Waiting for update"}
            2011 {$ES = "Waiting for user session reconnect"}
            2012 {$ES = "Waiting for user logoff"}
            2013 {$ES = "Waiting for user logon"}
            2014 {$ES = "Waiting To Install"}
            2015 {$ES = "Waiting Retry"}
            2016 {$ES = "Waiting For Presentation Mode"}
            2017 {$ES = "Waiting For Orchestration"}
            2018 {$ES = "Waiting For Network"}
            2019 {$ES = "Pending App-V Virtual Environment Update"}
            2020 {$ES = "Updating App-V Virtual Environment"}
            3000 {$ES = "Requirements not met"}
            3001 {$ES = "Host Platform Not Applicable"}
            4000 {$ES = "Unknown"}
            5000 {$ES = "Deployment failed"}
            5001 {$ES = "Evaluation failed"}
            5002 {$ES = "Deployment failed"}
            5003 {$ES = "Failed to locate content"}
            5004 {$ES = "Dependency installation failed"}
            5005 {$ES = "Failed to download dependent content"}
            5006 {$ES = "Conflicts with another application deployment"}
            5007 {$ES = "Waiting Retry"}
            5008 {$ES = "Failed to uninstall superseded deployment type"}
            5009 {$ES = "Failed to download superseded deployment type"}
            5010 {$ES = "Failed to updating App-V Virtual Environment"}
            default {$ES = $A.EnforcementState}
        }
        $A | Add-Member -NotePropertyName EnforcementState -NotePropertyValue $ES -Force
        if($AppDeploymentErrorAssetDetails.ErrorCode){
            $ErrorCode = "0x$([String]::Format("{0:x8}", $AppDeploymentErrorAssetDetails.ErrorCode))"
            switch ($ErrorCode){
                "0x87D00213" {$ErrorCode = "$ErrorCode Timeout occured"}
                "0x87D00324" {$ErrorCode = "$ErrorCode The software package was not detected after instillation"}
                "0x87D00325" {$ErrorCode = "$ErrorCode The software package uninstalled successfully, but a software detection rule was still found"}
                "0x87D00607" {$ErrorCode = "$ErrorCode Unable to Download Software"}
                "0x87d01201" {$ErrorCode = "$ErrorCode Unable to download install because hard drive is full"}
                "0x80091007" {$ErrorCode = "$ErrorCode Hash value incorrect"}
                "0x80041001" {$ErrorCode = "$ErrorCode SMS Agent Host service not running (ccmexec.exe)"}
                "0x87D01106" {$ErrorCode = "$ErrorCode Content is not available on Distribution Point"}
                "0x87D00443" {$ErrorCode = "$ErrorCode Cannot install because a conflicting process is running"}
                "0x80040154" {$ErrorCode = "$ErrorCode Class not registered"}
                "0x80041006" {$ErrorCode = "$ErrorCode Out of memory"}
                "0x800706ba" {$ErrorCode = "$ErrorCode The RPC server is unavalible"}
                "0x800706be" {$ErrorCode = "$ErrorCode The remote procedure call failed."}
                "0x80070057" {$ErrorCode = "$ErrorCode The parameter is incorrect."}
                "0x87d00231" {$ErrorCode = "$ErrorCode Transient error"}
                "0x80041033" {$ErrorCode = "$ErrorCode Shutting down"}
                "0x87d00317" {$ErrorCode = "$ErrorCode Unknown Error -2016410857"}
                "0x800705b4" {$ErrorCode = "$ErrorCode The Timeout period expired"}
                "0x87d0027c" {$ErrorCode = "$ErrorCode CI documents download timed out"}
                "0x00000652" {$ErrorCode = "$ErrorCode Another installation is already in progress. Complete that installation before proceeding with this install"}
                "0x80041009" {$ErrorCode = "$ErrorCode Not Available"}
                "0x00000643" {$ErrorCode = "$ErrorCode Dependency installation failed"}
                "0x87d00314" {$ErrorCode = "$ErrorCode Evaluation failed"}
                default {$ES = $ErrorCode}
            }
        }
        else{
            $ErrorCode = ""
        }
        $A | Add-Member -NotePropertyName ErrorCode -NotePropertyValue $ErrorCode -Force
        $A
    }
} 

Call-RCMServiceGui