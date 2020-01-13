Import-Module -Name Selenium

$time = Get-Date -Format "ddMMyyyy-HHmmss"

$moodleLoginUrl = # Moodle login page url here
$moodleCourseUrl = # Moodle Course page with Prezi links
$transcriptFile = ".\transcript - $($time).txt"

$username = Read-Host -Prompt "Enter your username: "
$password = Read-Host -Prompt "Enter your password: "

$Driver = Start-SeChrome

if ($Driver) {
    Enter-SeUrl -Driver $Driver -Url $moodleLoginUrl
    
    $emailField = Find-SeElement -Id "username" -Driver $Driver 
   
    Send-SeKeys -Element $emailField -Keys "$($username)@student.aau.dk"
    $pwField = Find-SeElement -Id "password" -Driver $Driver
    Send-SeKeys -Element $pwField -Keys $password
    Send-SeKeys -Element $pwField -Keys ([OpenQA.Selenium.Keys]::Enter)

    Enter-SeUrl -Driver $Driver -Url $moodleCourseUrl

    $allPreziAnchors = Find-SeElement -TagName a -Driver $Driver | Where-Object {$_.GetAttribute("href") -like "*prezi*"}
    
    $preziPages = @()

    foreach ($preziPage in $allPreziAnchors) {
        $preziPages += $preziPage.GetAttribute("href")
    }

    ForEach ($prezi in $preziPages) {
        Enter-SeUrl -Driver $Driver -Url $prezi
        $cookieBtn = Find-SeElement -Driver $Driver -ClassName "accept-cookie-container"
        if ($cookieBtn) {
            try {
            Invoke-SeClick -Driver $Driver -Element $cookieBtn
            } catch {<# Intentionally left blank #>}
        }
        $btn = Find-SeElement -Driver $Driver -Id "show-btn"
        $Driver.ExecuteScript("arguments[0].scrollIntoView(true);", $btn)
        Invoke-SeClick -Driver $Driver -Element $btn
        $transcript = Find-SeElement -Id "transcript-full-text" -Driver $Driver
        $transcript.Text | Out-File -FilePath transcript.txt -Append -Encoding utf8 
    }

    Stop-SeDriver -Driver $Driver
}