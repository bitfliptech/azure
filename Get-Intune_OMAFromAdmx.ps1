#Requires -Version 7.0

<#
    .SYNOPSIS
    Parses .admx files for use in Microsoft Endpoint Management Policies

    .DESCRIPTION
    Parses .admx files and will output either a PSCustomObject[] or an Excel Spreadsheet.
    The corresponding .adml file is used to bring in the Help for policies defined.
    For policy values that have an enumeration, then samples of the possible values are generated.
    If a policy does not have an enumeration of values, then a simple <enabled/> is provided. Depending on your use, you might also need <disabled/>    
    
    OMA-URI Format
    ./{user or device}/Vendor/MSFT/Policy/Config/{AreaName}/{PolicyName}

    Note: Can probably run with an earlier version than Powershell Core 7.x. Script was developed on a Mac and therefore it can be run cross-platform if needed

    .PARAMETER AdmxPath
    Path to .admx file. The required .adml file can be either in the same folder as the .admx file or in a subfolder according to the -DefaultLang parameter.

    .PARAMETER Excel
    Specifies if Excel spreadsheet should be generated.  If parameter is provided, the .xlsx file will be generated in the same folder as the .admx file
    Note: Requires ImportExcel Module (https://www.powershellgallery.com/packages/ImportExcel)

    .PARAMETER DefaultLang
    Language folder name containing the .adml file. The default without providing a value is 'en-us'.

    .PARAMETER Dev
    Display Debug Output

    .OUTPUTS
    If the -Excel parameter is not provided, then the output is a Powershell PSCustomObject[] that can used in a pipeline 

    .EXAMPLE
    .\Get-Intune_OMAFromAdmx.ps1 -AdmxPath  .\admx\ReaderADMTemplate\AcrobatReader2020.admx

    .EXAMPLE
    .\Get-Intune_OMAFromAdmx.ps1 -AdmxPath  .\admx\ReaderADMTemplate\AcrobatReader2020.admx | Out-GridView

    .EXAMPLE
    .\Get-Intune_OMAFromAdmx.ps1 -AdmxPath  .\admx\ReaderADMTemplate\AcrobatReader2020.admx -Dev -Excel
    DEBUG: Excel Output Requested
    DEBUG: Excel File Path: .\admx\ReaderADMTemplate\AcrobatReader2020.xlsx
    DEBUG: Generating Excel File

    .LINK
    Powershell Gallery ImportExcel Module: https://www.powershellgallery.com/packages/ImportExcel

    .NOTES   
    Name: Get-Intune_OMAFromAdmx.ps1
    Author: Nick Bowen
    DateCreated: 10.23.2020
#>

param (
    [parameter(Mandatory=$true)][String]$AdmxPath,
    [parameter(Mandatory=$false)][switch]$Excel,
    [parameter(Mandatory=$false)][String]$DefaultLang="en-us",
    [parameter(Mandatory=$false)][switch]$Dev
)

if ($Dev) {
	$DebugPreference = "Continue"	# Write-Debug statements will be written to console
}

if ($Excel) {
    Write-Debug "Excel Output Requested"
    if (-not [bool](Get-InstalledModule -Name ImportExcel -ErrorAction SilentlyContinue)) {
        Write-Debug "ImportExcel Module Missing"
        Write-Output "Please install the following module: https://www.powershellgallery.com/packages/ImportExcel"
        exit 1
    }
} else {
    Import-Module ImportExcel
}

# Platform Dependent Path Separator (ie. \ or /)
$PathSep = [IO.Path]::DirectorySeparatorChar

if (-not (Test-Path -Path $AdmxPath)) {
    Write-Output "ADMX File Missing ($AdmxPath)"
    exit 1
}

$AdmxFileObj = Get-ChildItem -Path $AdmxPath
$AdmxFile = $AdmxFileObj.FullName
$AdmlName = $AdmxFileObj.Name -replace '\.admx$','.adml'
$AdmlDir = $AdmxFileObj.Directory.FullName
$AdmlFile = "$AdmlDir$PathSep$AdmlName"
if (-not (Test-Path -Path $AdmlFile)) {
    $AdmlDirAlt = "$AdmlDir$PathSep$DefaultLang"
    $AdmlFileAlt = "$AdmlDirAlt$PathSep$AdmlName"
    if (-not (Test-Path -Path $AdmlFileAlt)) {
        Write-Output "ADML File Missing.  Should be located in either: `n$AdmlFile `n$AdmlFileAlt)"
        exit 1
    } else {
        $Local:AdmlFile = $AdmlFileAlt
    }
}
if ($Excel) {
    $Local:ExcelName = $AdmxFileObj.Name -replace '\.admx$','.xlsx'
    $Local:ExcelPath = "$AdmlDir$PathSep$ExcelName"
    Write-Debug "Excel File Path: $ExcelPath"
}

function Get-AreaName {
    [CmdletBinding()]
    param (
        [parameter(mandatory=$true)][System.Xml.XmlDocument]$ADMX    
    )
    begin {}
    process {
        # The {AreaName} format is {AppName}~{SettingType}~{CategoryPathFromAdmx}
        try {
            $AppName = $ADMX.policyDefinitions.policyNamespaces.target.prefix
            $Categories = @{}           
            $ADMX.policyDefinitions.categories.category | ForEach-Object {
                # Using child::node() was more reliable than using parentCategory for -XPath
                #   if ([bool](Select-Xml -Xml $_ -XPath 'child::node()')) {
                if((Select-Xml -Xml $_ -XPath 'child::node()').Node.Name -eq 'parentCategory') {
                    $Categories.Add($_.name,$_.parentCategory.ref)
                } else {                    
                    $Categories.Add($_.name,$_.name)
                }
            }
            $ParentCategories = @()
            $Categories.GetEnumerator() | ForEach-Object {
                # Test for Parent Categories
                if (-not $Categories.ContainsKey($_.Value)) {
                    $ParentCategories += $_.Key
                }
            }
            $ParentCategories | ForEach-Object { $Categories[$_] = $_ }
            $Keys = @()
            $Categories.Keys | ForEach-Object { $Keys += $_ }
            foreach ($Key in $Keys) {    
                if ($Categories[$Key] -eq $Key) {
                    $Categories[$Key] = "$AppName~Policy~$Key"
                } else {
                    $Categories[$Key] = "$AppName~Policy~$($Categories[$Key])~$Key"
                }
            }           
            $Categories
        } catch {
            $PositionStr = $_.InvocationInfo.PositionMessage -replace '\r\n\+.*','' -replace '\n\+.*',''
            $ExceptionClean = $_.Exception.Message -replace '\r\n\+.*','' -replace '\n\+.*',''
            Write-Debug $PositionStr
            Write-Debug $ExceptionClean
        }       
    }
    end {}
}

function Get-Policies {
    [CmdletBinding()]
    param (
        [parameter(mandatory=$true)][System.Xml.XmlDocument]$ADMX,
        [parameter(mandatory=$true)][System.Xml.XmlDocument]$ADML,
        [parameter(mandatory=$true)][Hashtable]$AreaName
    )
    begin {}
    process {
        # The {AreaName} format is {AppName}~{SettingType}~{CategoryPathFromAdmx}
        try {
            $Policies = @()
            $ADMX.policyDefinitions.policies.policy | ForEach-Object {
                $PolicyDetails = $_
                $Help = ""
                $Value = "<enabled/>"
                if ($null -ne (Select-Xml -Xml $PolicyDetails -XPath '@explainText')) {
                    $ADLMID = $PolicyDetails.explainText -replace '.*string\.','' -replace '\)|\(',''
                    $HelpTemp = ($ADML.policyDefinitionResources.resources.stringTable.string | Where-Object { $_.id -eq $ADLMID }).'#text'
                    $Local:Help = $HelpTemp -replace ',',' '
                }                   
                if ([bool](Select-Xml -Xml $PolicyDetails -XPath 'elements/enum')) {                    
                    $ValueTemp = ""
                    if([bool](Select-Xml -Xml $PolicyDetails -XPath 'elements/enum/item/value/decimal')) {
                        $PolicyDetails.elements.enum.item.value.decimal | ForEach-Object {  
                            $Local:ValueTemp += "`n<data id=`"$($PolicyDetails.elements.enum.valueName)`" value=`"$([int]$_.value)`"/>"
                        }                        
                    } elseif ([bool](Select-Xml -Xml $PolicyDetails -XPath 'elements/enum/item/value/string')) {
                        $PolicyDetails.elements.enum.item.value.string | ForEach-Object {                            
                            $Local:ValueTemp += "`n<data id=`"$($PolicyDetails.elements.enum.valueName)`" value=`"$($_)`"/>"
                        }                           
                    }
                    $Local:Value += $ValueTemp
                }
                Switch -regex ($_.class) {
                    'User|Both' {
                        $PolicyObj = [PSCustomObject]@{
                            name   = $PolicyDetails.name
                            omauri = "./User/Vendor/MSFT/Policy/Config/$($AreaName[$PolicyDetails.parentCategory.ref])/$($PolicyDetails.name)"
                            value  = $value
                            help   = $Help
                            scope  = "user"
                        }
                        $Policies += $PolicyObj
                    }
                    'Device|Machine|Both' {
                        $PolicyObj = [PSCustomObject]@{
                            name   = $PolicyDetails.name
                            omauri = "./Device/Vendor/MSFT/Policy/Config/$($AreaName[$PolicyDetails.parentCategory.ref])/$($PolicyDetails.name)"
                            value  = $value
                            help   = $Help
                            scope  = "device"
                        }
                        $Policies += $PolicyObj
                    }
                }
                # <enabled/>	<data id=""BrowserSignin"" value=""0""/>                
            }
            $Policies
        } catch {
            $PositionStr = $_.InvocationInfo.PositionMessage -replace '\r\n\+.*','' -replace '\n\+.*',''
            $ExceptionClean = $_.Exception.Message -replace '\r\n\+.*','' -replace '\n\+.*',''
            Write-Debug $PositionStr
            Write-Debug $ExceptionClean
        }       
    }
    end {}
}

try {
    $ADMX = [xml](Get-Content -Path $AdmxFile)
    $ADML = [xml](Get-Content -Path $AdmlFile)
    
    $AreaName = Get-AreaName -ADMX $ADMX
    $AllPolicies = Get-Policies -ADMX $ADMX -ADML $ADML -AreaName $AreaName
    
    if ($Excel) {
        if (Test-Path $ExcelPath) {
            Write-Debug "Removing Existing Excel File: $ExcelPath"
            Remove-Item $ExcelPath -Force
        }
        
        Write-Debug "Generating Excel File"
        $ExcelObj = $AllPolicies | Where-Object { $_.scope -eq 'user' } | Export-Excel -Path $ExcelPath -WorksheetName "User" -AutoSize -AutoFilter -PassThru
        Set-Format -Address $ExcelObj.Workbook.Worksheets["User"].Cells -WrapText -VerticalAlignment Top
        Close-ExcelPackage $ExcelObj 

        $ExcelObj = $AllPolicies | Where-Object { $_.scope -eq 'device' } | Export-Excel -Path $ExcelPath -WorksheetName "Device" -AutoSize -AutoFilter -PassThru
        Set-Format -Address $ExcelObj.Workbook.Worksheets["Device"].Cells -WrapText -VerticalAlignment Top
        Close-ExcelPackage $ExcelObj 
    } else {
        $AllPolicies
    }
} catch {
    $PositionStr = $_.InvocationInfo.PositionMessage -replace '\r\n\+.*','' -replace '\n\+.*',''
    $ExceptionClean = $_.Exception.Message -replace '\r\n\+.*','' -replace '\n\+.*',''
    Write-Debug $PositionStr
    Write-Debug $ExceptionClean
}