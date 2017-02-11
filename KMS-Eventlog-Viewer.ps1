#Get-Command -Module kds

#Get-KdsConfiguration
#Get-KdsRootKey -Verbose

$cCounter = 1
$Report = @()

#$After = Get-Date -Year 2017 -Month 1 -Day 20 -Hour 10 -Minute 10 -Second 00 -Millisecond 000
#$After = (Get-Date).AddHours(-1)
$After =  (Get-Date).AddDays(-1)
#$After = (Get-Date).AddDays(-7)
#$After = (Get-Date).AddYears(-1)
#TODO : https://msdn.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/creating-a-graphical-date-picker
#       https://technet.microsoft.com/en-us/library/ff730942.aspx

Function getLocalTime($UTCTime)
{
    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
    Return $LocalTime
}

$(foreach ($entry in (Get-EventLog -Logname "Key Management Service" -After $After)) {
    #Write-Output $entry.ReplacementStrings[3]
    if ($cCounter -eq 72) {
        Write-Host "."
        $cCounter = 1
    } else {
        Write-Host "." -NoNewline
        $cCounter++
    }
	$ReportRow = "" | select Result,Client,DateTime,ClientGuid,LicenseSku,<#NeededToActivate,#>IsVM,LicenseState,MinutesToStateExpiration
    $ReportRow.Client = $entry.ReplacementStrings[3]
    $ReportRow.DateTime = getLocalTime((Get-Date $entry.ReplacementStrings[5]))
    $ReportRow.ClientGuid = $entry.ReplacementStrings[4]
    switch ($entry.ReplacementStrings[9])
    {
        #region WindowsServer2016
        '8c1c5410-9f39-4805-8c9d-63a07706358f' {$ReportRow.LicenseSku = "Windows Server 2016 Standard"}
        #endregion
        #region WindowsServer2012R2
        '00091344-1ea4-4f37-b789-01750ba6988c' {$ReportRow.LicenseSku = "Windows Server 2012 R2 Datacenter"}
        'b3ca044e-a358-4d68-9883-aaa2941aca99' {$ReportRow.LicenseSku = "Windows Server 2012 R2 Standard"}
        'b743a2be-68d4-4dd3-af32-92425b7bb623' {$ReportRow.LicenseSku = "Windows Server 2012 R2 Cloud Storage"}
        '21db6ba4-9a7b-4a14-9e29-64a60c59301d' {$ReportRow.LicenseSku = "Windows Server Essentials 2012 R2"}
        #endregion
        #region WindowsServer2012
        'd3643d60-0c42-412d-a7d6-52e6635327f6' {$ReportRow.LicenseSku = "Windows Server 2012 Datacenter"}
        'f0f5ec41-0d55-4732-af02-440a44a3cf0f' {$ReportRow.LicenseSku = "Windows Server 2012 Standard"}
        '95fd1c83-7df5-494a-be8b-1300e1c9d1cd' {$ReportRow.LicenseSku = "Windows Server 2012 MultiPoint Premium"}
        '7d5486c7-e120-4771-b7f1-7b56c6d3170c' {$ReportRow.LicenseSku = "Windows Server 2012 MultiPoint Standard"}
        #endregion
        #region WindowsServer2008R2
        '68531fb9-5511-4989-97be-d11a0f55633f' {$ReportRow.LicenseSku = "Windows Server 2008 R2 Standard"}
        '7482e61b-c589-4b7f-8ecc-46d455ac3b87' {$ReportRow.LicenseSku = "Windows Server 2008 R2 Datacenter"}
        '620e2b3d-09e7-42fd-802a-17a13652fe7a' {$ReportRow.LicenseSku = "Windows Server 2008 R2 Enterprise"}
        'a78b8bd9-8017-4df5-b86a-09f756affa7c' {$ReportRow.LicenseSku = "Windows Server 2008 R2 Web"}
        'cda18cf3-c196-46ad-b289-60c072869994' {$ReportRow.LicenseSku = "Windows Server 2008 R2 ComputerCluster"}
        #endregion
        #region WindowsServer2008
        'ad2542d4-9154-4c6d-8a44-30f11ee96989' {$ReportRow.LicenseSku = "Windows Server 2008 Standard"}
        '2401e3d0-c50a-4b58-87b2-7e794b7d2607' {$ReportRow.LicenseSku = "Windows Server 2008 StandardV"}
        '68b6e220-cf09-466b-92d3-45cd964b9509' {$ReportRow.LicenseSku = "Windows Server 2008 Datacenter"}
        'fd09ef77-5647-4eff-809c-af2b64659a45' {$ReportRow.LicenseSku = "Windows Server 2008 DatacenterV"}
        'c1af4d90-d1bc-44ca-85d4-003ba33db3b9' {$ReportRow.LicenseSku = "Windows Server 2008 Enterprise"}
        '8198490a-add0-47b2-b3ba-316b12d647b4' {$ReportRow.LicenseSku = "Windows Server 2008 EnterpriseV"}
        'ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b' {$ReportRow.LicenseSku = "Windows Server 2008 Web"}
        '7afb1156-2c1d-40fc-b260-aab7442b62fe' {$ReportRow.LicenseSku = "Windows Server 2008 ComputerCluster"}
        #endregion

        #region Windows10
        #endregion
        #region Windows8.1
        '81671aaf-79d1-4eb1-b004-8cbbe173afea' {$ReportRow.LicenseSku = "Windows 8.1 Enterprise"}
        '113e705c-fa49-48a4-beea-7dd879b46b14' {$ReportRow.LicenseSku = "Windows 8.1 EnterpriseN"}
        '096ce63d-4fac-48a9-82a9-61ae9e800e5f' {$ReportRow.LicenseSku = "Windows 8.1 Professional WMC"}
        'c06b6981-d7fd-4a35-b7b4-054742b7af67' {$ReportRow.LicenseSku = "Windows 8.1 Professional"}
        '7476d79f-8e48-49b4-ab63-4d0b813a16e4' {$ReportRow.LicenseSku = "Windows 8.1 ProfessionalN"}
        'fe1c3238-432a-43a1-8e25-97e7d1ef10f3' {$ReportRow.LicenseSku = "Windows 8.1 Core"}
        '78558a64-dc19-43fe-a0d0-8075b2a370a3' {$ReportRow.LicenseSku = "Windows 8.1 CoreN"}
        #endregion
        #region Windows8
        'a00018a3-f20f-4632-bf7c-8daa5351c914' {$ReportRow.LicenseSku = "Windows 8 Professional WMC"}
        'a98bcd6d-5343-4603-8afe-5908e4611112' {$ReportRow.LicenseSku = "Windows 8 Professional"}
        'ebf245c1-29a8-4daf-9cb1-38dfc608a8c8' {$ReportRow.LicenseSku = "Windows 8 ProfessionalN"}
        '458e1bec-837a-45f6-b9d5-925ed5d299de' {$ReportRow.LicenseSku = "Windows 8 Enterprise"}
        'e14997e7-800a-4cf7-ad10-de4b45b578db' {$ReportRow.LicenseSku = "Windows 8 EnterpriseN"}
        'c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60' {$ReportRow.LicenseSku = "Windows 8 Core"}
        '197390a0-65f6-4a95-bdc4-55d58a3b0253' {$ReportRow.LicenseSku = "Windows 8 CoreN"}
        #endregion
        #region Windows7
        'ae2ee509-1b34-41c0-acb7-6d4650168915' {$ReportRow.LicenseSku = "Windows 7 Enterprise"}
        '1cb6d605-11b3-4e14-bb30-da91c8e3983a' {$ReportRow.LicenseSku = "Windows 7 EnterpriseN"}
        'b92e9980-b9d5-4821-9c94-140f632f6312' {$ReportRow.LicenseSku = "Windows 7 Professional"}
        '54a09a0d-d57b-4c10-8b69-a842d6590ad5' {$ReportRow.LicenseSku = "Windows 7 ProfessionalN"}
        #endregion
        #region WindowsVista
        'cfd8ff08-c0d7-452b-9f60-ef5c70c32094' {$ReportRow.LicenseSku = "Windows Vista Enterprise"}
        'd4f54950-26f2-4fb4-ba21-ffab16afcade' {$ReportRow.LicenseSku = "Windows Vista EnterpriseN"}
        '4f3d1606-3fea-4c01-be3c-8d671c401e3b' {$ReportRow.LicenseSku = "Windows Vista Business"}
        '2c682dc2-8b68-4f63-a165-ae291d4cf138' {$ReportRow.LicenseSku = "Windows Vista BusinessN"}
        #endregion

        #region Office2016
        'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' {$ReportRow.LicenseSku = "Project 2016 Standard"}
        'dedfa23d-6ed1-45a6-85dc-63cae0546de6' {$ReportRow.LicenseSku = "Office 2016 Standard"}
        '83e04ee1-fa8d-436d-8994-d31a862cab77' {$ReportRow.LicenseSku = "Skype for Business 2016"}
        'aa2a7821-1827-4c2c-8f1d-4513a34dda97' {$ReportRow.LicenseSku = "Visio 2016 Standard"}
        #endregion
        #region Office2013
        '6ee7622c-18d8-4005-9fb7-92db644a279b' {$ReportRow.LicenseSku = "Office Access 2013"}
        'f7461d52-7c2b-43b2-8744-ea958e0bd09a' {$ReportRow.LicenseSku = "Office Excel 2013"}
        'a30b8040-d68a-423f-b0b5-9ce292ea5a8f' {$ReportRow.LicenseSku = "Office InfoPath 2013"}
        '1b9f11e3-c85c-4e1b-bb29-879ad2c909e3' {$ReportRow.LicenseSku = "Office Lync 2013"}
        'dc981c6b-fc8e-420f-aa43-f8f33e5c0923' {$ReportRow.LicenseSku = "Office Mondo 2013"}
        'efe1f3e6-aea2-4144-a208-32aa872b6545' {$ReportRow.LicenseSku = "Office OneNote 2013"}
        '771c3afa-50c5-443f-b151-ff2546d863a0' {$ReportRow.LicenseSku = "Office OutLook 2013"}
        '8c762649-97d1-4953-ad27-b7e2c25b972e' {$ReportRow.LicenseSku = "Office PowerPoint 2013"}
        '4a5d124a-e620-44ba-b6ff-658961b33b9a' {$ReportRow.LicenseSku = "Office Project Pro 2013"}
        '427a28d1-d17c-4abf-b717-32c780ba6f07' {$ReportRow.LicenseSku = "Office Project Standard 2013"}
        '00c79ff1-6850-443d-bf61-71cde0de305f' {$ReportRow.LicenseSku = "Office Publisher 2013"}
        #'b13afb38-cd79-4ae5-9f7f-eed058d750ca' {$ReportRow.LicenseSku = "Office Visio Standard 2013"}
        'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' {$ReportRow.LicenseSku = "Office Visio Standard 2013"}
        'e13ac10e-75d0-4aff-a0cd-764982cf541c' {$ReportRow.LicenseSku = "Office Visio Pro 2013"}
        'd9f5b1c6-5386-495a-88f9-9ad6b41ac9b3' {$ReportRow.LicenseSku = "Office Word 2013"}
        'b322da9c-a2e2-4058-9e4e-f59a6970bd69' {$ReportRow.LicenseSku = "Office Professional Plus 2013"}
        'b13afb38-cd79-4ae5-9f7f-eed058d750ca' {$ReportRow.LicenseSku = "Office Standard 2013"}
        #endregion
        #region Office2010
        '8ce7e872-188c-4b98-9d90-f8f90b7aad02' {$ReportRow.LicenseSku = "Office Access 2010"}
        'cee5d470-6e3b-4fcc-8c2b-d17428568a9f' {$ReportRow.LicenseSku = "Office Excel 2010"}
        '8947d0b8-c33b-43e1-8c56-9b674c052832' {$ReportRow.LicenseSku = "Office Groove 2010"}
        'ca6b6639-4ad6-40ae-a575-14dee07f6430' {$ReportRow.LicenseSku = "Office InfoPath 2010"}
        '09ed9640-f020-400a-acd8-d7d867dfd9c2' {$ReportRow.LicenseSku = "Office Mondo 2010"}
        'ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c' {$ReportRow.LicenseSku = "Office Mondo 2010"}
        'ab586f5c-5256-4632-962f-fefd8b49e6f4' {$ReportRow.LicenseSku = "Office OneNote 2010"}
        'ecb7c192-73ab-4ded-acf4-2399b095d0cc' {$ReportRow.LicenseSku = "Office OutLook 2010"}
        '45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a' {$ReportRow.LicenseSku = "Office PowerPoint 2010"}
        'df133ff7-bf14-4f95-afe3-7b48e7e331ef' {$ReportRow.LicenseSku = "Office Project Pro 2010"}
        '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' {$ReportRow.LicenseSku = "Office Project Standard 2010"}
        'b50c4f75-599b-43e8-8dcd-1081a7967241' {$ReportRow.LicenseSku = "Office Publisher 2010"}
        '92236105-bb67-494f-94c7-7f7a607929bd' {$ReportRow.LicenseSku = "Office Visio Premium 2010"}
        'e558389c-83c3-4b29-adfe-5e4d7f46c358' {$ReportRow.LicenseSku = "Office Visio Pro 2010"}
        '9ed833ff-4f92-4f36-b370-8683a4f13275' {$ReportRow.LicenseSku = "Office Visio Standard 2010"}
        '2d0882e7-a4e7-423b-8ccc-70d91e0158b1' {$ReportRow.LicenseSku = "Office Word 2010"}
        '6f327760-8c5c-417c-9b61-836a98287e0c' {$ReportRow.LicenseSku = "Office Professional Plus 2010"}
        '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' {$ReportRow.LicenseSku = "Office Standard 2010"}
        'ea509e87-07a1-4a45-9edc-eba5a39f36af' {$ReportRow.LicenseSku = "Office Small Business Basics 2010"}
        #endregion

        #'' {$ReportRow.LicenseSku = ""}
        Default {$ReportRow.LicenseSku = $entry.ReplacementStrings[9]}
    }
    #$ReportRow.NeededToActivate = $entry.ReplacementStrings[2]
    $ReportRow.Result = $entry.ReplacementStrings[1]
    #$ReportRow.IsVM = $entry.ReplacementStrings[6]
    switch ($entry.ReplacementStrings[6])
    {
        0 {$ReportRow.IsVM = "No"}
        1 {$ReportRow.IsVM = "Yes"}
        Default {$ReportRow.IsVM = "Unknown"}
    }
    #$ReportRow.LicenseState = $entry.ReplacementStrings[7]
    switch ($entry.ReplacementStrings[7])
    {
        0 {$ReportRow.LicenseState = "Unlicensed"}
        1 {$ReportRow.LicenseState = "Activated"}
        2 {$ReportRow.LicenseState = "Grace Period"}
        3 {$ReportRow.LicenseState = "Out-of-Tolerance Grace Period"}
        4 {$ReportRow.LicenseState = "Non-Genuine Grace Period"}
        5 {$ReportRow.LicenseState = "Notifications Mode"}
        6 {$ReportRow.LicenseState = "Extended Grace Period"}
        Default {$ReportRow.LicenseState = "Unknown"}
    }
    $ReportRow.MinutesToStateExpiration = $entry.ReplacementStrings[8]
    $Report += $ReportRow

 }) 
 
  $Report | Select LicenseSku -Unique
 
 $Report | Out-GridView -Title "KMS Eventlog entries since $($After.ToString())" -Wait
