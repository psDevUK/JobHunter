<#

.Synopsis

Automates applying for jobs on totaljobs.com

.DESCRIPTION

Uses Selenium and Powershell to automate the applying for various jobs on totaljobs.com this will only automatically apply for jobs that do not require you to go to the external company website

.EXAMPLE

Find-TotalJobs -SeleniumSupportFilePath $env:SeleniumSupport -SeleniumDriverFilePath $Env:SeleniumDriver -SeleniumPath $env:SeleniumPath -Email $env:emailaddress -Password $env:totaljobsPassword -JobTitle "Vmware" -JobType 'Work From Home' -JobSalary '50,000'

.EXAMPLE

Find-TotalJobs -SeleniumSupportFilePath $env:SeleniumSupport -SeleniumDriverFilePath $Env:SeleniumDriver -SeleniumPath $env:SeleniumPath -Email $env:emailaddress -Password $env:totaljobsPassword -JobTitle "Infrastructure Analyst" -Location 'Hampshire'

.NOTES

You need to have Selenium downloaded and configured prior to using this function as it relies on Selenium to do the hard work for you

.COMPONENT

The component this cmdlet belongs to Adam Bacon

.FUNCTIONALITY

The functionality that best describes this cmdlet is the ability to automatically apply for numerous jobs without you having to do it

#>

function Find-TotalJobs {

[CmdletBinding()]

param

(

# Enter the full path to where you have the WebDriver.Support.dll for Selenium on your system

[Parameter(Mandatory = $true, Position = 0)]

[ValidateNotNullOrEmpty()]

[ValidateScript({ if (-not (Test-Path $_)) { throw "$_ is not a valid path to the dll file needed" }return $true })]

[System.IO.FileInfo]$SeleniumSupportFilePath,



# Enter the full path to where you have the WebDriver.dll for Selenium on your system

[Parameter(Mandatory = $true, Position = 1)]

[ValidateNotNullOrEmpty()]

[ValidateScript({ if (-not (Test-Path $_)) { throw "$_ is not a valid path to the dll file needed" }return $true })]

[System.IO.FileInfo]$SeleniumDriverFilePath,



# Enter the full path for the system path to the Selenium folder containing the msedgedriver.exe

[Parameter(Mandatory = $true, Position = 2)]

[ValidateNotNullOrEmpty()]

[ValidateScript({ if (-not (Test-Path $_)) { throw "$_ is not a valid path to the dll file needed" }return $true })]

[System.IO.FileInfo]$SeleniumPath,



# Enter your email address that you have registered with on TotalJobs.com

[Parameter(Mandatory = $true, Position = 3)]

[ValidateScript({ if ($_ -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw "'$_' is not a valid Email address" }return $true })]

[string]$Email,



# Enter your password for the TotalJobs website

[Parameter(Mandatory = $true, Position = 4)]

[ValidateNotNullOrEmpty()]

[String]$Password,



# Enter the job title of the job you are looking for such as "Infrastructure Analyst"

[Parameter(Mandatory = $true, Position = 5)]

[ValidateNotNullOrEmpty()]

[String]$JobTitle,



# Enter the job location you are looking to obtain a job in, such as Devon

[Parameter(Position = 6)]

[String]$Location,



# Optional selecting the type of work you are looking for

[Parameter(Position = 7)]

[ValidateSet("Permanent", "Work From Home", "Contract", "Part Time")]

[String]$JobType,



# Optional to select to return jobs that meet a particular salary

[Parameter(Position = 8)]

[ValidateSet("30,000", "40,000", "50,000", "60,000", "70,000", "80,000")]

[String]$JobSalary,



# Optional to select the start page that the searches begin on, this is from the pagnation along the bottom this is defaulted to 1 but can be changed

[Parameter(Position = 9)]

[int]$StartPage = 1,



# Optional to select the end page that the searches end on, this is from the pagnation along the bottom this is defaulted to 7 but can be changed

[Parameter(Position = 9)]

[int]$EndPage = 7

)



begin {

# Load Selenium DLLs

Add-Type -Path $SeleniumSupportFilePath

Add-Type -Path $SeleniumDriverFilePath



# Setup Edge driver

$edgeOptions = New-Object OpenQA.Selenium.Edge.EdgeOptions

$service = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService("$SeleniumPath")

$driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($service, $edgeOptions)



# Create WebDriverWait (10 seconds timeout)

$wait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($driver, [TimeSpan]::FromSeconds(10))



# Step 1: Navigate to jobsite page and accept cookies

$driver.Navigate().GoToUrl("https://www.totaljobs.com/account/signin")



try {

$cookieAcceptButton = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::Id("ccmgt_explicit_accept"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch {

return $null

}

return $null

})

if ($cookieAcceptButton) {

$cookieAcceptButton.Click()

Start-Sleep -Seconds 1

}

}

catch { }

# Enter email

$emailField = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("input[data-testid='email-input']"))

if ($el.Displayed) { return $el }

}

catch { return $null }

return $null

})

$emailField.Clear()

$emailField.SendKeys("$Email")



# Enter password

$passwordField = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("input[data-testid='password-input']"))

if ($el.Displayed) { return $el }

}

catch { return $null }

return $null

})

$passwordField.Clear()

$passwordField.SendKeys("$Password")



# Click login button

$signInButton = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("button[data-testid='login-submit-btn']"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch { return $null }

return $null

})

$signInButton.Click()

Start-Sleep -Seconds 3



# Search Field

$searchField = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("input[data-at='searchbar-keyword-input']"))

if ($el.Displayed) { return $el }

}

catch { return $null }

return $null

})

$searchField.Clear()

$searchField.SendKeys("$JobTitle")



if ($null -ne $Location) {

# Location Field

$searchField = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("input[data-at='searchbar-location-input']"))

if ($el.Displayed) { return $el }

}

catch { return $null }

return $null

})

$searchField.Clear()

$searchField.SendKeys("$Location")

}



# Click Search button

$searchButton = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("button[data-at='searchbar-search-button']"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch { return $null }

return $null

})

if ($searchButton) { Write-Host "Button found, clicking..." } else { Write-Host "Button not found." }

$searchButton.Click()





if ($null -ne $JobType) {

$wfhElement = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::XPath("//a[@data-at='facet-link' and normalize-space(text())='$JobType']"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch { return $null }

return $null

})



# Click the link if found

if ($wfhElement) {

Write-Host "Clicking 'Work From Home' link..."

$wfhElement.Click()

}

else {

Write-Host "Link not found or not clickable."

}

}

if ($null -ne $JobSalary) {

# Wait for the salary filter link to appear by data-at attribute

$salaryLink = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::XPath("//a[@data-at='facet-link' and normalize-space(text())='at least £$($JobSalary)']"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch { return $null }

return $null

})



if ($salaryLink) {

Write-Host "Clicking 'at least £$JobSalary' salary filter..."

$salaryLink.Click()

}

else {

Write-Host "Salary filter not found."

}

}

#Page should now only contain the matching jobs for the criteria selected



}

process {

# Get the new URL

$currentUrl = $driver.Url

Write-Host "Navigated to: $currentUrl"



# Now you can use this as your base for pagination

$page = $StartPage

$maxPages = $EndPage



while ($page -lt $maxPages) {

$pagedUrl = "$($currentUrl)?page=$page"

$driver.Navigate().GoToUrl($pagedUrl)

# Get all job links

$jobLinks = $driver.FindElements([OpenQA.Selenium.By]::CssSelector("a[href*='/job/']"))

# Store the original tab

$originalTab = $driver.CurrentWindowHandle

foreach ($link in $jobLinks) {

$href = $link.GetAttribute("href")

if ($href -and $href.StartsWith("https://www.totaljobs.com/job/")) {

Write-Host "Opening: $href"

# Open job tab

$driver.ExecuteScript("window.open(arguments[0], '_blank');", $href)

Start-Sleep -Seconds 1



# Switch to job tab

$jobTab = $driver.WindowHandles[-1]

$driver.SwitchTo().Window($jobTab)



try {

$applyButton = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("button[data-testid='harmonised-apply-button']"))

if ($el.Displayed) { return $el }

}

catch { return $null }

return $null

})



if ($applyButton) {

$buttonText = $applyButton.Text.Trim()

$isDisabled = -not $applyButton.Enabled



if ($isDisabled -and $buttonText -match "Already applied") {

Write-Host "Already applied — skipping job."

$driver.Close() # Close job tab

$driver.SwitchTo().Window($originalTab)

}

else {

Write-Host "Button found, clicking..."

$handlesBefore = $driver.WindowHandles

$applyButton.Click()

Start-Sleep -Seconds 3



# Detect new tab

$handlesAfter = $driver.WindowHandles

$newTabs = $handlesAfter | Where-Object { $handlesBefore -notcontains $_ }



if ($newTabs.Count -gt 0) {

Write-Host "New tab detected — switching and closing..."

$driver.SwitchTo().Window($newTabs[0])

Start-Sleep -Seconds 1

$driver.Close() # Close new tab

$driver.SwitchTo().Window($jobTab)

$driver.Close() # Close job tab

$driver.SwitchTo().Window($originalTab)

continue # Skip rest of logic

}



# Proceed with sendApplication

$nextButton = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]] {

param($d)

try {

$el = $d.FindElement([OpenQA.Selenium.By]::CssSelector("button[data-testid='sendApplication']"))

if ($el.Displayed -and $el.Enabled) { return $el }

}

catch { return $null }

return $null

})



if ($nextButton) {

Write-Host "Clicking 'Send Application'..."

$nextButton.Click()

Start-Sleep -Seconds 6

}

else {

Write-Host "No follow-up button found — skipping job."

}



$driver.Close() # Close job tab

$driver.SwitchTo().Window($originalTab)

}

}

}

catch {

Write-Host "Error finding apply button: $_"

$driver.Close()

$driver.SwitchTo().Window($originalTab)

}

}

}

$page++

}

}

end {

Write-Host "Script Finished..."

}

}
