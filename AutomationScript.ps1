Param(
    $workbook = "test.pbix",

    [switch]
    $publish =$True
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\UIAutomation.0.8.7B3.NET35\UIAutomation.dll"

function TestUnknownMessageDialog(){
    #Test if Message dialog exists
    try  {
        $dialog = $window | Get-UiaChildWindow | Read-UiaControlAutomationId
    }
    catch {
    }
    if ($dialog -match "Challenge") {
        $message = ($window | Get-UiaChildWindow | Get-UiaText | Read-UiaControlName)
        Throw ("Unknown message dialog opened. " + $message)

    }    
    return $false
}



function TestLoadDialog(){
    $exist = ($window  | Test-UiaControlState -SearchCriteria @{automationID="KoLoadToReportDialog"})    
    return  $exist
}

function TestStatusDialog(){
    $exist = ($window  | Test-UiaControlState -SearchCriteria @{automationID="StatusDialog"} )    
    return  $exist
}


function WaitStatusDialog(){
    $null = TestUnknownMessageDialog
    while(TestStatusDialog){
        $null = TestUnknownMessageDialog
        sleep -Seconds 5
        "*"
    }
}

function WaitLoadDialog(){
    $null = TestUnknownMessageDialog
    while(TestLoadDialog){
        #$null = TestUnknownMessageDialog
        $text = $null
        $text += $window | Get-UiaChildWindow | Get-UiaText | Read-UiaControlName
        if ($text -match "error"){
            Write-Error "Error. $text"
        } else {
            sleep -Seconds 5
            "*"
        }
    }
}

function Quit(){
    $null = $window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::LMENU)
    $null = $window.Keyboard.TypeText("FX")
}


########################
##        BODY

$ready = $null
$workbookName = (get-item $workbook).BaseName
$window = $null
$statisticsStart = Get-Date
$statisticsEnd = $null

#Open Power BI workbook
Invoke-Item $workbook

#kill process if the workbook is opened
try {
    $processId = Get-UiaWindow -ProcessName PBIDesktop -Name "$workbookName*" | Read-UiaControlProcessId    
    "Power BI workbook $workbookName is already open. Killing process..."
    Kill -Id $processId
} catch {
}

#Wait for main window of Power BI Desktop    
"Waiting for Power BI to open"
while (!$ready){
    try{
        $ready = Get-UiaWindow  -ProcessName PBIDesktop -Name "$workbookName - Power BI Desktop" | test-UiaControlState -SearchCriteria @{name="Refresh"} 
        if (!$ready){
            sleep -Seconds 5  
        }
    }catch{"*"}
}
"Power BI opened"

#Pick Power BI window
$window = Get-UiaWindow  -ProcessName PBIDesktop -Name "$workbookName - Power BI Desktop" 


#Press REFRESH
"Refreshing"
if (!($window | Test-UiaControlState -SearchCriteria @{name="Refresh"})){
    Write-Error "Not able to press refresh"
}
$null = $window | Get-UiaButton -Name "Refresh" | Invoke-UiaButtonClick
    
#Wait to complete
WaitLoadDialog
"Data refreshed locally"

#Press Save
"New Saving"
if (!($window | Test-UiaControlState -SearchCriteria @{name="Save"})){
    Write-Error "Not able to save file"
}
$null = $window | Get-UiaButton -Name "Save" | Invoke-UiaButtonClick
#Wait to complete
WaitStatusDialog


#Publish

if($publish){
"New Publishing"
if (!($window | Test-UiaControlState -SearchCriteria @{name="Publish"})){
    Write-Error "Not able to press publish"
}
$null = $window | Get-UiaButton -Name "Publish" | Invoke-UiaButtonClick
#Wait to complete
WaitStatusDialog
"Searching Dashboard"
if (!($window | Test-UiaControlState -SearchCriteria @{name="My workspace"} )){
    Write-Error "Not able to select Workspace"
}
"Selecting Dashboard"
$null = $window | Get-UiaText  -Name "My workspace" |Invoke-UiaButtonClick
#WaitStatusDialog
if (!($window | Test-UiaControlState -SearchCriteria @{name="Select"})){
    Write-Error "Not able to select Workspace"
}
$null = $window | Get-UiaButton -Name "Select" | Invoke-UiaButtonClick
#WaitStatusDialog
WaitLoadDialog
 #replace existing report on Power BI web
 "Replace Dashboard"
   if($window | Test-UiaControlState -SearchCriteria @{name="Replace"}){
       $null = $window | Get-UiaButton -Name "Replace" | Invoke-UiaButtonClick
       "Replacing dashboard on web."
    }

     $published = $false
   while(!$published){
       $published = $window | Test-UiaControlState -SearchCriteria @{name="Success!"}
      if (!$published) {
           sleep -Seconds 2
      }
   }


    if ($published){
     "Published"
     
      $null = $window | Get-UiaChildWindow | Get-UiaButton -Name "Close" | Invoke-UiaButtonClick
    
        #Quit
      # Quit
   }else{
       #set unknown result
      $message = "Unknown result of publishing. "
       #set result from last window
     $message += ($window | Get-UiaChildWindow -automationID MessageDialog | Get-UiaText | Read-UiaControlName)      
      Throw $message
   }
} else {
   Quit
 }
 $statisticsEnd = Get-Date
"Total seconds {0:N0}" -f ($statisticsEnd - $statisticsStart).TotalSeconds



