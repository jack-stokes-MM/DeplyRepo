#Azure variables
$ClientId = Get-AutomationVariable -Name 'Azure Automation ClientID'
$TenantId = Get-AutomationVariable -Name 'Azure Automation TenantID'
$Cert = Get-AutomationCertificate -Name 'Azure Automation'
$CertificateThumbprint = $Cert.thumbprint
$Tenant = 'forthepeople0.onmicrosoft.com'

#Culture AMP variables
$caClientId = Get-AutomationVariable -Name 'Culture Amp Suvey - Client ID' 
$caClientSecret = Get-AutomationVariable -Name 'Culture Amp Suvey - Client Secret'
$CaBaseUrl = "https://api.cultureamp.com/v1"
$SurveyId = "392213d1-d8ba-4f7c-be3b-7f8c5775f668"

# Azure Storage account details
$StorageAccountName = "saehremployeesurveys"
$StorageAccountRG = "saehremployeesurveys"

#Template file variables
$DestinationFolder = "$env:temp"
$SourceFolder = "$env:temp"
$TemplateFileName = "ExitSurveyTemplate.docx"
$ContainerName = "templates"

#SharePoint Variables
$SharePointURL = "https://forthepeople0.sharepoint.com"
$SharePointSite = "$SharePointURL/teams/HREmployeeSurveys"
$SharePointListName = "M&M Exit Survey Sentiment"

#Open AI and sentiment variables
$Keywords = @("Hostile", "Harassment", "discrimination", "retaliation")
$OpenAIKey = Get-AutomationVariable -Name 'OpenAIKey'

#Email variables:
$fromEmail = "employeesurveys@forthepeople.com"
$TriggerDay = 1
#Force parameter forces email to send every run
$force = $false

# Define recipients
$recipientEmails = @(
    "racheljohnson@forthepeople.com"
)

$bccRecipients = @(
    "HREmployeeSurveyReveiwers@forthepeople.com"
)

#Used for weekly HR email
$hrTeamEmails = @(
    "racheljohnson@forthepeople.com"
)

try {
    Connect-PnPOnline -Url $SharePointSite -Tenant $Tenant -ClientId $ClientId -Thumbprint $Cert.Thumbprint
    Write-Output "Connected to PnP via Service Principal"
}
catch {
    Write-Output "An error occurred connecting to SharePoint: $_"
    Throw 
}

function Get-CultureAmpToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret
    )
    
    $CultureTokenUrl = "https://api.cultureamp.com/v1/oauth2/token"
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $CaClientId
        client_secret = $CaClientSecret
        scope         = "target-entity:8ed17dce-9eca-4383-a9e1-54f82c362b6d:surveys-read,employees-read,employee-demographics-read"
    }

    try {
        $Response = Invoke-RestMethod -Uri $CultureTokenUrl -Method Post -Body $Body
        return $Response.access_token
    }
    catch {
        Write-Error "Failed to obtain token: $($_.Exception.Message)"
        return $null
    }
}
function Get-CultureAmpHeaders {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    return @{
        "Authorization" = "Bearer $CultureToken"
        "Accept"        = "application/json"
    }
}

function Get-CultureAmpSurveyResponses {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SurveyId   
    )

    $ResponsesUrl = "$CaBaseUrl/surveys/$SurveyId/responses"
    try {
        $SurveyResponses = @()
        $Pagination = $ResponsesUrl
        while ($Pagination) {
            $Responses = Invoke-RestMethod -Uri $Pagination -Method Get -Headers $CultureHeaders
            $SurveyResponses += $Responses.data
            if ($Responses.pagination.nextPath) {
                $Pagination = "https://api.cultureamp.com$($Responses.pagination.nextPath)"
            }
            else { 
                $Pagination = $null 
            }
        }
        return $SurveyResponses
    }
    catch {
        Write-Error "Failed to create distribution: $($_.Exception.Message)"
        return $null
    }
}

function Get-EmployeeData {
    param (        
        [Parameter(Mandatory = $true)]
        [string]$EmployeeId            
    )

    $Employee = (Invoke-RestMethod -Uri "$CaBaseUrl/employees/$EmployeeId" -Method Get -Headers $CultureHeaders).data
    $EmployeeDemo = (Invoke-RestMethod -Uri "$CaBaseUrl/employees/$($Employee.id)/demographics" -Method Get -Headers $CultureHeaders).data
    
    foreach ($EmpDemo in $EmployeeDemo) { 
        $Employee | Add-Member -MemberType NoteProperty -Name $EmpDemo.name -Value $EmpDemo.value
    }

    if ($Employee) {
        return $Employee
    }
    else {
        return "Employee not found."
    }
}

function Get-SurveyQuestions {
    # Get questions
    [array]$SurveyQuestions = (Invoke-RestMethod -Uri "https://api.cultureamp.com/v1/surveys/$($SurveyId)/questions" -Method Get -Headers $CultureHeaders).data
    return $SurveyQuestions
}

function Get-ExistingResponses {
    $items = Get-PnPListItem -List $SharePointListName -PageSize 50000
    $ResponseIDs = $items | ForEach-Object { $_.FieldValues.SurveyID }
    return $ResponseIDs
}
function Upload-PDFsToSharePoint {
    param (
        [string]$FileLocation,
        [string]$ParentOffice,
        [datetime]$SubmittedAt
    )

    try {
        # Connect to SharePoint
        Write-Host "Connecting to SharePoint site: $SiteURL"
        $PdfFiles = Get-ChildItem -Path $FileLocation 
        
        if ($PdfFiles.Count -eq 0) {
            Write-Host "No PDF files found in $LocalFolderPath"
            return
        }
        Write-Host "Found $($PdfFiles.Count) PDF files to upload" -ForegroundColor Yellow

        # Construct the target folder path
        $MonthName = (Get-Date($SubmittedAt)).ToString("MMMM")
        $Year = (Get-Date).Year
        $TargetPath = "Shared Documents/Survey Results/Exit Surveys/$MonthName $Year/$ParentOffice"
        
        # Check if folder exists, create if it doesn't
        $Folder = Get-PnPFolder -Url $TargetPath -ErrorAction SilentlyContinue 
        if (!$Folder) {
            Write-Host "Creating folder structure: $TargetPath"
            $Folder = Resolve-PnPFolder -SiteRelativePath $TargetPath
        }

        # Upload each PDF
        $SuccessCount = 0
        $ErrorCount = 0
        $Results = @()
        
        foreach ($Pdf in $PdfFiles) {
            try {
                Write-Host "Uploading: $($Pdf.Name)..." -NoNewline
                
                # Check if file already exists
                $ExistingFile = Get-PnPFile -Url "$TargetPath/$($Pdf.Name)" -ErrorAction SilentlyContinue
                if ($ExistingFile) {
                    Write-Host "File already exists, updating..." -NoNewline
                }
                
                # Upload the file
                $File = Add-PnPFile -Path $Pdf.FullName -Folder $TargetPath -ErrorAction Stop
                
                Write-Host "Success!" -ForegroundColor Green
                $SuccessCount++
                
                # Add to results array
                $Results += [PSCustomObject]@{
                    FileName = $Pdf.Name
                    Status   = "Success"
                    URL      = $SharePointURL + $($File.ServerRelativeUrl)
                }
            }
            catch {
                Write-Host "Failed!" -ForegroundColor Red
                Write-Host "Error: $_" -ForegroundColor Red
                $ErrorCount++
                
                # Add to results array
                $Results += [PSCustomObject]@{
                    FileName = $Pdf.Name
                    Status   = "Failed"
                    Error    = $_.Exception.Message
                }
            }
        }

        # Summary
        Write-Host "`nUpload Summary:" -ForegroundColor Yellow
        Write-Host "Successfully uploaded: $SuccessCount" -ForegroundColor Green
        Write-Host "Failed uploads: $ErrorCount" -ForegroundColor $(if ($ErrorCount -eq 0) { "Green" } else { "Red" })
        Remove-Item -Path $PdfFiles -Force
        
        return $Results
    }
    catch {
        Write-Host "Error in Upload-PDFsToSharePoint: $_" -ForegroundColor Red
        throw
    }
    finally {
        Write-Host ""
    }
}

function Create-PDF {
    param ([object]$Data)
    
    $Word = $null
    $Doc = $null
    $Range = $null
    
    try {
        $Word = New-Object -ComObject word.application 
        $Word.visible = $false 
        $TempDocFile = "$DestinationFolder\$($Data.'name') - Exit SurveyTemp10.docx"
        $PdfPath = "$DestinationFolder\$($Data.'name') - Exit Survey.pdf"
        
        Copy-Item $TemplateFile.fullname $TempDocFile -Force
        $Doc = $Word.documents.open($TempDocFile)

        # Process standard fields
        foreach ($Item in $Data.Keys) {
            if ($Item -in @("name", "office", "jobtitle", "department", "parentoffice", "manager", 
                    "get.hire.date", "submission.date.formatted", "s1.qual.score", "s2.qual.score", 
                    "s3.qual.score", "ov.qual.score")) {
                $Range = $Doc.Bookmarks.Item($Item.Replace(".", "_")).Range
                $Range.Text = "$($Data.$Item)"
                $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                $Range = $null
            }
        }

        # Process response data
        foreach ($Item in $Data.responseData.keys) {
            $Base = $Item.Replace(".", "_")
            $Bookmarks = $Doc.Bookmarks | select -ExpandProperty name
            
            if ($Data.responseData.$Item.comment -and ($Base + "_comment" -in $Bookmarks)) {
                $Range = $Doc.Bookmarks.Item($Base + "_comment").Range
                $Range.Text = $Data.responseData.$Item.comment
                $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                $Range = $null
            }
            if ($Data.responseData.$Item.score) {
                if ($Base + "_score" -in $Bookmarks) {
                    $Range = $Doc.Bookmarks.Item($Base + "_score").Range
                    $Range.Text = $Data.responseData.$Item.score.ToString()
                    $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                    $Range = $null
                }
            }
            if ($Data.responseData.$Item.'additional.comment') {
                if ($Base + "_comment" -in $Bookmarks) {
                    $Range = $Doc.Bookmarks.Item($Base + "_comment").Range
                    $Range.Text = "$($Data.responseData.$Item.'additional.comment')"
                    $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                    $Range = $null
                }
            }
            foreach ($Option in 1..3) {
                if ($Data.responseData.$Item."Option$Option") {
                    $Range = $Doc.Bookmarks.Item($Base + "_Option$Option").Range
                    $Range.Text = "$($Data.responseData.$Item."option$Option")"
                    $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                    $Range = $null
                }
            }
        }

        # Set default values for empty bookmarks
        foreach ($Bookmark in $Doc.Bookmarks) {
            if ($Bookmark.name -like "*_comment") {
                $Range = $Doc.Bookmarks.Item($Bookmark.name).Range
                $Range.Text = "No Comment Submitted" 
                $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                $Range = $null
            }
            if ($Bookmark.name -like "*_option*") {
                $Range = $Doc.Bookmarks.Item($Bookmark.name).Range
                $Range.Text = "" 
                $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
                $Range = $null
            }
        }

        $Doc.saveas([ref]$PdfPath, [ref]17)
        return $PdfPath
    }
    catch {
        Write-Error "PDF Creation Error: $_"
        return $null
    }
    finally {
        if ($Doc) {
            $Doc.Close($false)
            $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Doc)
            $Doc = $null
        }
        if ($Word) {
            $Word.Quit()
            $var = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
            $Word = $null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 500  # Add delay
        if (Test-Path $TempDocFile) { 
            Remove-Item -Path $TempDocFile -Force -ErrorAction SilentlyContinue 
        }
    }
}
function Get-SentimentAnalysis {
    param (
        [object]$Data,
        [object]$Keywords
    )

    $OpenAIHeaders = @{
        'Authorization' = "Bearer $OpenAIKey"
        'Content-Type'  = 'application/json'
    }
    
    $SentimentObj = @()
    $QuestionCodes = $Data.responseData.Keys
    
    foreach ($Question in $QuestionCodes) {
        $SelectionText = @()
        if ($Data.responseData.$Question.option1) {
            if ($Data.responseData.$Question.option1) { $SelectionText += "1.$($Data.responseData.$Question.option1)" }
            if ($Data.responseData.$Question.option2) { $SelectionText += " 2.$($Data.responseData.$Question.option2)" }
            if ($Data.responseData.$Question.option3) { $SelectionText += " 3.$($Data.responseData.$Question.option3)" }
            
            if (!$($Data.responseData.$Question.comment)) {
                $Comment = "No comment submitted"
            }
            else {
                $Comment = $($Data.responseData.$Question.comment)
            }
            
            $SentimentObj += @{
                "$Question" = @( "Question:$($Data.responseData.$Question.Question)", "Selections:$($SelectionText)", "Employee Comment: $($Comment)")
            }
        }
        elseif ($Data.responseData.$Question.comment) {
            $SentimentObj += @{
                "$Question" = @( "Question:$($Data.responseData.$Question.Question)", "Employee Comment:$($Data.responseData.$Question.comment)")
            }
        }
        elseif ($Data.responseData.$Question.additional_comment) {
            $SentimentObj += @{
                "$Question" = @( "Question:$($Data.responseData.$Question.Question)", "Employee additional Comment:$($Data.responseData.$Question.additional_comment)")
            }
        }
    }

    $Keywords = foreach ($Keyword in $Keywords) { '"' + $($Keyword) + '";' }
    $Text = "$($SentimentObj | ConvertTo-Json)"
    
    $Content = @"
   You are an AI trained to analyze survey responses. Analyze the provided responses and return a JSON object with the following structure:

   {
       "overall_sentiment": string,          // Must be exactly one of: "positive", "neutral", "negative"
       "inappropriate_content": boolean,      // true if potentially inappropriate content is detected
       "inappropriate_content_details": string, // Description of inappropriate content if found, null string if none. Do not use employee names if mentioned. 
       "detected_keywords": {                // Check for presence of specific keywords. 
           [keyword: string]: boolean        // Each keyword from [$($Keywords)] as key with boolean value if true. null if no key words
       }
   }

   Requirements:
   1. Response must be valid JSON only - no prefix or additional text
   2. Sentiment must be strictly "positive", "neutral", or "negative"
   3. Check responses for keywords listed in $($Keywords) case-insensitively
   4. Flag inappropriate content that may be concerning or require HR review
   5. Include question in inappropriate_content_questionid array only when inappropriate content is found

    CONCERNING CONTENT DEFINITION:
    1. Compliance or legal risks
    2. Discrimination, harassment, or hostile workplace claims
    3. Potential safety violations
    4. Retaliation claims
    5. Mentions of sensitive personal information
    6. Inflammatory language directed at specific individuals

    OUTPUT FORMAT:
    Return a JSON object with EXACTLY these fields:
    {
        "overall_sentiment": string,       // MUST be "positive", "neutral", or "negative" only
        "inappropriate_content": boolean,  // true if concerning content is detected
        "inappropriate_content_details": string, // Brief description if inappropriate_content is true, otherwise ""
        "contains_keywords": {
            // For each keyword in the list, indicate if present (case-insensitive)
        }
    }
"@

    $Body = @{
        model       = "gpt-4o-mini"
        messages    = @(
            @{
                role    = "system"
                content = $Content
            },
            @{
                role    = "user"
                content = $Text
            }
        )
        temperature = 0.3
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Headers $OpenAIHeaders -Method Post -Body $Body
        return $Response.choices[0].message.content.Trim()
    }
    catch {
        Write-Error "Failed to analyze sentiment: $_"
        return "ERROR"
    }
}

function Log-RecordAndSentiment {
    param (
        [string]$JsonData,
        [string]$SurveyId,
        [string]$EmployeeName,
        [string]$EmployeeId,
        [string]$SubmittedAt,
        [string]$SurveyURL
    )

    try {
        $JsonObject = $JsonData | ConvertFrom-Json
    }
    catch { 
        return $JsonData
    }

    $Keywords = @()
    foreach ($Key in $JsonObject.contains_keywords.PSObject.Properties.Name) {
        if ($JsonObject.contains_keywords.$Key -eq $true) {
            $Keywords += $Key
        }
    }
    if (!$Keywords) {
        $Keywords += "None"
    }
    
    $KeywordsString = $Keywords -join ", "
    $ListName = $SharePointListName

    $ListItem = @{
        "Title"                       = "$EmployeeName"
        "EmployeeID"                  = $EmployeeId
        "SurveyID"                    = $SurveyId
        "OverallSentiment"            = $JsonObject.overall_sentiment
        "InappropriateContent"        = $JsonObject.inappropriate_content
        "InappropriateContentDetails" = $JsonObject.inappropriate_content_details
        "DetectedKeywords"            = $KeywordsString
        "SubmittedAt"                 = $SubmittedAt
        "SurveyURL"                   = $SurveyURL
    }

    Add-PnPListItem -List $ListName -Values $ListItem
    Write-Output "Item logged successfully to SharePoint List: $ListName"
}
# Function to send email using Graph API
function Send-GraphEmail {
    param (
        [Parameter(Mandatory = $true)]
        [string] $From,
        
        [Parameter(Mandatory = $true)]
        [string[]] $To,

        [Parameter(Mandatory = $false)]
        [string[]] $Bcc = @(),
        
        [Parameter(Mandatory = $false)]
        [switch] $Force,
        
        [Parameter(Mandatory = $false)]
        [string] $StorageAccountName = $StorageAccountName,
        
        [Parameter(Mandatory = $false)]
        [string] $ResourceGroupName = $StorageAccountRG,
        
        [Parameter(Mandatory = $false)]
        [string] $TableName = "MonthlyEmailSentStatus"
    )
    try {
        
        try {
            # Connect to Azure using the service principal
            Connect-AzAccount -ServicePrincipal -TenantId $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertificateThumbprint
            Write-Output "Connected to AzAccount via Service Principal"
        }
        catch {
            Write-Output "An error occurred connecting to AZ: $_"
            Throw 
        }
        
        # Connect to Azure and get token for Graph API
        $AccessToken = Get-AzAccessToken -ResourceTypeName "MSGraph"
        if (-not $AccessToken -or -not $AccessToken.Token) {
            throw "Failed to acquire Microsoft Graph access token"
        }
            
        # Get current date information
        $today = Get-Date
        $currentDay = $today.Day
        $currentDayOfWeek = $today.DayOfWeek
        
        # Determine target month for survey results
        # If we're in the first 7 days of the month, use previous month's folder
        $targetMonth = if ($currentDay -le 7) { $today.AddMonths(-1) } else { $today }
        
        # Logic: Only proceed if it's the trigger day or later, AND it's a weekday (or forced)
        $isWeekday = $currentDayOfWeek -ne 'Saturday' -and $currentDayOfWeek -ne 'Sunday'
        $isTriggerDayOrAfter = $currentDay -ge $TriggerDay
        
        if (-not $Force) {
            if (-not $isTriggerDayOrAfter) {
                Write-Output "Not yet trigger day ($TriggerDay). Use -Force to override."
                return @{
                    Success    = $false
                    Reason     = "Not yet trigger day"
                    EmailsSent = 0
                }
            }
            
            if (-not $isWeekday) {
                Write-Output "Today is $currentDayOfWeek. Waiting for next weekday. Use -Force to override."
                return @{
                    Success    = $false
                    Reason     = "Not a weekday"
                    EmailsSent = 0
                }
            }
            
            # Check if this is the first weekday on or after the trigger day by checking the Azure Storage Table
            try {
                # Get storage account using RBAC instead of keys
                $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName
                
                # Get storage context using the current authenticated session
                $ctx = $storageAccount.Context
                
                # Ensure table exists
                $storageTable = Get-AzStorageTable -Name $TableName -Context $ctx -ErrorAction SilentlyContinue
                if (-not $storageTable) {
                    Write-Output "Creating table $TableName..."
                    $storageTable = New-AzStorageTable -Name $TableName -Context $ctx
                }
                
                # Access the table
                $cloudTable = (Get-AzStorageTable -Name $TableName -Context $ctx).CloudTable
                
                # Create a unique key for this month's email (e.g., "ExitSurvey-2025-03")
                $partitionKey = "ExitSurvey"
                $rowKey = $targetMonth.ToString("yyyy-MM")
                
                # Check if we already sent an email for this month
                $operation = Get-AzTableRow -table $cloudTable -partitionKey $partitionKey -rowKey $rowKey -ErrorAction SilentlyContinue
                
                if ($operation) {
                    Write-Output "Email already sent for $($targetMonth.ToString('MMMM yyyy')) on $($operation.SentDate). Use -Force to send again."
                    return @{
                        Success    = $false
                        Reason     = "Already sent this month on $($operation.SentDate)"
                        EmailsSent = 0
                    }
                }
            }
            catch {
                Write-Output "Warning: Could not check email sent status in Azure Table: $_"
                Write-Output "Proceeding with sending email..."
            }
        }
        
        # Format folder name properly (e.g., "January 2025")
        $folderName = $targetMonth.ToString("MMMM yyyy")
        $folderPath = "Shared Documents/Survey Results/Exit Surveys/$folderName"
        
        # Construct email subject with month information
        $monthName = $targetMonth.ToString("MMMM")
        $Subject = "$monthName Exit Survey Results"
        
        # Get folder URL
        $folderUrl = "$SharePointSite/$folderPath"
        Write-Output "Targeting folder: $folderUrl"
        Write-Output "Sending email to $($To.Count) recipients..."
        
        # Create recipients array
        $Recipients = @()
        foreach ($Recipient in $To) {
            $Recipients += @{
                "emailAddress" = @{
                    "address" = $Recipient
                }
            }
        }
        # Create BCC recipients array if provided
        $BccRecipients = @()
        if ($Bcc.Count -gt 0) {
            foreach ($BccRecipient in $Bcc) {
                $BccRecipients += @{
                    "emailAddress" = @{
                        "address" = $BccRecipient
                    }
                }
            }
        }
        
        # Create HTML body with the SharePoint link
        $htmlBody = @"
<p>Good morning,</p>
<p>This is to inform you that last month's exit survey results are ready to be viewed! Please <a href='$folderUrl'>click here</a> to view the most recent results.</p>
<p>Please let us know if you have any questions.</p>
<p>Thank you,</p>
<p>HR</p>
"@
        
        # Create message payload
        $Message = @{
            "message" = @{
                "subject"      = $Subject
                "body"         = @{
                    "contentType" = "html"
                    "content"     = $htmlBody
                }
                "toRecipients" = $Recipients
                "from"         = @{
                    "emailAddress" = @{
                        "address" = $From
                    }
                }
            }
        }

        # Add BCC recipients if provided
        if ($BccRecipients.Count -gt 0) {
            $Message.message.bccRecipients = $BccRecipients
        }
        
        # Send email using Graph API
        $GraphEndpoint = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        $Headers = @{
            "Authorization" = "Bearer $($AccessToken.Token)"
            "Content-Type"  = "application/json"
        }
        
        $Response = Invoke-RestMethod -Uri $GraphEndpoint -Headers $Headers -Method POST -Body ($Message | ConvertTo-Json -Depth 10)
        
        # Record sent status in Azure Table
        if (-not $Force -and $cloudTable) {
            try {
                $properties = @{
                    PartitionKey = $partitionKey
                    RowKey       = $rowKey
                    SentDate     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    Recipients   = ($To -join ", ")
                    TargetFolder = $folderUrl
                }
                
                Add-AzTableRow -table $cloudTable -property $properties -partitionKey $partitionKey -rowKey $rowKey -ErrorAction Stop
                Write-Output "Recorded email sent status in Azure Table Storage"
            }
            catch {
                Write-Output "Warning: Could not record email sent status: $_"
            }
        }
        
        return @{
            Success      = $true
            EmailsSent   = $To.Count
            TargetFolder = $folderUrl
        }
    }
    catch {
        Write-Error "Failed to send email: $_"
        throw $_
    }
}

function Send-WeeklySentimentEmail {
    param (
        [Parameter(Mandatory = $true)]
        [string] $From,
        
        [Parameter(Mandatory = $true)]
        [string[]] $To,
        
        [Parameter(Mandatory = $false)]
        [switch] $Force,
        
        [Parameter(Mandatory = $false)]
        [string] $StorageAccountName = $StorageAccountName,
        
        [Parameter(Mandatory = $false)]
        [string] $ResourceGroupName = $StorageAccountRG,
        
        [Parameter(Mandatory = $false)]
        [string] $TableName = "WeeklySentimentEmailStatus",
        
        [Parameter(Mandatory = $false)]
        [string] $SharePointSiteURL = $SharePointSite,
        
        [Parameter(Mandatory = $false)]
        [string] $SentimentListName = $SharePointListName
    )
    try {
        # Make sure AzTable module is available
        if (-not (Get-Module -ListAvailable -Name AzTable)) {
            Write-Output "AzTable module not found. Installing..."
            Install-Module AzTable -Scope CurrentUser -Force -RequiredVersion 2.1.0
        }
        
        # Import the module
        Import-Module AzTable -RequiredVersion 2.1.0
        
        # Get current date information
        $today = Get-Date
        $currentDayOfWeek = $today.DayOfWeek
        
        # Only proceed if it's Monday (or forced)
        if (-not $Force -and $currentDayOfWeek -ne 'Monday') {
            Write-Output "Today is $currentDayOfWeek. The weekly sentiment email is only sent on Mondays. Use -Force to override."
            return @{
                Success    = $false
                Reason     = "Not Monday"
                EmailsSent = 0
            }
        }
        
        # Calculate the week start and end dates for the email subject
        # For PowerShell, Sunday is 0 and Monday is 1 in DayOfWeek.value__
        $weekStart = (Get-Date).Date.AddDays( - ((Get-Date).DayOfWeek.value__) + 1) # Monday of current week
        if ($currentDayOfWeek -eq 'Sunday') {
            # If today is Sunday, use the previous week
            $weekStart = $weekStart.AddDays(-7)
        }
        $weekEnd = $weekStart.AddDays(6) # Sunday
        
        # Check if we've already sent an email for this week
        if (-not $Force) {
            try {
                # Connect to Azure (if not already connected)
                try {
                    $context = Get-AzContext
                    if (-not $context) {
                        Connect-AzAccount -ServicePrincipal -TenantId $TenantId -ApplicationId $ClientId -CertificateThumbprint $CertificateThumbprint
                        Write-Output "Connected to AzAccount via Service Principal"
                    }
                }
                catch {
                    Write-Output "An error occurred connecting to AZ: $_"
                    throw
                }
                
                # Get storage account using RBAC
                $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName
                
                # Get storage context using the current authenticated session
                $ctx = $storageAccount.Context
                
                # Ensure table exists
                $storageTable = Get-AzStorageTable -Name $TableName -Context $ctx -ErrorAction SilentlyContinue
                if (-not $storageTable) {
                    Write-Output "Creating table $TableName..."
                    $storageTable = New-AzStorageTable -Name $TableName -Context $ctx
                }
                
                # Access the table
                $cloudTable = (Get-AzStorageTable -Name $TableName -Context $ctx).CloudTable
                
                # Create a unique key for this week's email (e.g., "WeeklySentiment-2025-03-03")
                $partitionKey = "WeeklySentiment"
                $rowKey = $weekStart.ToString("yyyy-MM-dd")
                
                # Check if we already sent an email for this week
                $operation = Get-AzTableRow -table $cloudTable -partitionKey $partitionKey -rowKey $rowKey -ErrorAction SilentlyContinue
                
                if ($operation) {
                    Write-Output "Weekly sentiment email already sent for week starting $rowKey on $($operation.SentDate). Use -Force to send again."
                    return @{
                        Success    = $false
                        Reason     = "Already sent this week on $($operation.SentDate)"
                        EmailsSent = 0
                    }
                }
            }
            catch {
                Write-Output "Warning: Could not check email sent status in Azure Table: $_"
                Write-Output "Proceeding with sending email..."
            }
        }
        
        # Make sure we have a connection to SharePoint
        try {
            $connection = Get-PnPConnection -ErrorAction SilentlyContinue
            if (-not $connection) {
                # Try to connect to SharePoint
                Write-Output "No active SharePoint connection found. Connecting to SharePoint..."
                Connect-PnPOnline -Url $SharePointSiteURL -Tenant $Tenant -ClientId $ClientId -Thumbprint $CertificateThumbprint
            }
        }
        catch {
            Write-Output "Error checking SharePoint connection: $_"
            # Attempt to connect anyway
            Connect-PnPOnline -Url $SharePointSiteURL -Tenant $Tenant -ClientId $ClientId -Thumbprint $CertificateThumbprint
        }
        
        # Get Graph API access token
        $AccessToken = Get-AzAccessToken -ResourceTypeName "MSGraph"
        if (-not $AccessToken -or -not $AccessToken.Token) {
            throw "Failed to acquire Microsoft Graph access token"
        }
        
        # Create subject line
        $Subject = "Weekly Exit Survey Sentiment Review - Week of $($weekStart.ToString('MMMM dd'))"
        
        # Create HR dashboard URL - make sure to use the right SharePoint URL
        $DashboardUrl = "$SharePointSiteURL/Lists/$SentimentListName/AllItems.aspx"
        $DashboardUrl = $DashboardUrl.Replace("&", "")

        # Generate HR insights using explicit date calculations
        $sevenDaysAgo = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
        
        try {
            # Count all surveys from past 7 days using explicit date
            $caml = "<View><Query><Where><Geq><FieldRef Name='SubmittedAt' /><Value Type='DateTime'>$sevenDaysAgo</Value></Geq></Where></Query></View>"
            $recentItems = Get-PnPListItem -List $SentimentListName -Query $caml
            $recentSentimentCount = $recentItems.Count
            
            # Count negative sentiment surveys
            $negativeCount = 0
            $inappropriateCount = 0
            
            foreach ($item in $recentItems) {
                # Check for negative sentiment
                if ($item.FieldValues.OverallSentiment -eq "negative") {
                    $negativeCount++
                }
                
                # Check for inappropriate content
                if ($item.FieldValues.InappropriateContent -eq $true) {
                    $inappropriateCount++
                }
            }
        }
        catch {
            Write-Output "Error retrieving sentiment data: $_"
            # Set fallback values
            $recentSentimentCount = "Unable to retrieve"
            $negativeCount = "Unable to retrieve"
            $inappropriateCount = "Unable to retrieve"
        }
        
        # Create recipients array
        $Recipients = @()
        foreach ($Recipient in $To) {
            $Recipients += @{
                "emailAddress" = @{
                    "address" = $Recipient
                }
            }
        }
        
        # Create HTML body with SharePoint link and insights
        $htmlBody = @"
<p>Hello Team,</p>
<p>Here is the weekly sentiment analysis review for exit surveys submitted during the week of $($weekStart.ToString('MMMM dd')) to $($weekEnd.ToString('MMMM dd')).</p>

<h3>Weekly Insights:</h3>
<ul>
    <li>New exit survey responses in the past week: <strong>$recentSentimentCount</strong></li>
    <li>Responses with negative sentiment: <strong>$negativeCount</strong></li>
    <li>Responses flagged with inappropriate content: <strong>$inappropriateCount</strong></li>
</ul>

<p>Please <a href='$DashboardUrl'>click here</a> to access the full sentiment analysis dashboard.</p>

<p>This is an automated message sent every Monday to help track and monitor employee exit surveys.</p>

<p>Thank you,</p>
<p>Modern Workplace Team</p>
"@
        
        # Create message payload
        $Message = @{
            "message" = @{
                "subject"      = $Subject
                "body"         = @{
                    "contentType" = "html"
                    "content"     = $htmlBody
                }
                "toRecipients" = $Recipients
                "from"         = @{
                    "emailAddress" = @{
                        "address" = $From
                    }
                }
            }
        }
        
        # Send email using Graph API
        $GraphEndpoint = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        $Headers = @{
            "Authorization" = "Bearer $($AccessToken.Token)"
            "Content-Type"  = "application/json"
        }
        
        $Response = Invoke-RestMethod -Uri $GraphEndpoint -Headers $Headers -Method POST -Body ($Message | ConvertTo-Json -Depth 10)
        
        # Record sent status in Azure Table
        if ((-not $Force) -and $cloudTable) {
            try {
                $properties = @{
                    SentDate           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    Recipients         = ($To -join ", ")
                    WeekStart          = $weekStart.ToString("yyyy-MM-dd")
                    WeekEnd            = $weekEnd.ToString("yyyy-MM-dd")
                    ResponseCount      = $recentSentimentCount
                    NegativeCount      = $negativeCount
                    InappropriateCount = $inappropriateCount
                }
                
                Add-AzTableRow -table $cloudTable -property $properties -partitionKey $partitionKey -rowKey $rowKey -ErrorAction Stop
                Write-Output "Recorded weekly email sent status in Azure Table Storage"
            }
            catch {
                Write-Output "Warning: Could not record email sent status: $_"
            }
        }
        
        return @{
            Success       = $true
            EmailsSent    = $To.Count
            WeekStart     = $weekStart.ToString("yyyy-MM-dd")
            WeekEnd       = $weekEnd.ToString("yyyy-MM-dd")
            ResponseCount = $recentSentimentCount
        }
    }
    catch {
        Write-Error "Failed to send weekly sentiment email: $_"
        throw $_
    }
}

function Get-TemplateFile {
    try {
        return Get-ChildItem -Path "$($SourceFolder)\$TemplateFileName" -ErrorAction Stop
    }
    catch {
        # Copy file from storage account to source folder
        Write-Host "Template file not found locally. Downloading from Azure Storage..."
        
        # Get storage account using RBAC instead of keys
        $StorageAccount = Get-AzStorageAccount -ResourceGroupName $StorageAccountRG -Name $StorageAccountName
        
        # Get context from the authenticated account
        $StorageContext = $StorageAccount.Context
        
        # Download the file
        Get-AzStorageBlobContent -Container $ContainerName -Blob $TemplateFileName -Destination "$SourceFolder\$TemplateFileName" -Context $StorageContext -Force
        
        # Return the file info
        return Get-ChildItem -Path "$($SourceFolder)\$TemplateFileName"
    }
}
# Main Script Execution
# Call the function and store the result
$TemplateFile = Get-TemplateFile

# Use the template file
if ($TemplateFile) {
    Write-Host "Template file found/downloaded successfully: $($TemplateFile.FullName)"
}
else {
    Write-Host "Failed to retrieve template file" -ForegroundColor Red
    $TemplateFile
}

#Get culture AMP information
$FirstDay = Get-Date -Date $(Get-Date) -Day 1
$CultureToken = Get-CultureAmpToken -ClientId $CaClientId -ClientSecret $CaClientSecret
$CultureHeaders = Get-CultureAmpHeaders -Token $CultureToken
$SurveyResponses = Get-CultureAmpSurveyResponses -SurveyId $SurveyId | Where-Object { $_.submittedAt -ge $FirstDay }
$SurveyQuestions = Get-SurveyQuestions | Where-Object { $_.status -eq "active" } 

#Gets all recorded responsed
$existingResponses = Get-ExistingResponses
$Employees = @{}
$SentimentObj = @{}
foreach ($SurveyResponse in $SurveyResponses) {
    #Process survey if response ID isn't already recorded
    if (-not ($SurveyResponse.id -in $existingResponses) ) {
        $ThisObject = @{}
        $EmployeeData = $(Get-EmployeeData -EmployeeId $SurveyResponse.employeeId)
        
        $ThisObject += @{
            "name"                      = $EmployeeData.name
            "responseid"                = $SurveyResponse.id
            "jobtitle"                  = $EmployeeData.'Job Title'
            "department"                = $EmployeeData.Department
            "parentoffice"              = $EmployeeData.'Parent Office'
            "office"                    = $EmployeeData.Office
            "manager"                   = $EmployeeData.Manager
            "get.hire.date"             = $EmployeeData.startDate
            "submission.date.formatted" = $SurveyResponse.submittedAt.ToString("yyyy-M-dd")
        }
        
        #creates Response data property on object
        $ThisObject.Add("responseData", @{})
        #parses through each expected survey question and saves responses
        foreach ($SurveyQuestion in $SurveyQuestions) {
            $QuestionData = $SurveyQuestions | Where-Object { $_.code -eq $SurveyQuestion.code }
            $ResponseData = $SurveyResponse.answers | Where-Object { $_.questionId -eq $QuestionData.id }
            
            $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                "Question" = $QuestionData.description.en
            }
            
            if ($ResponseData.textvalue) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "comment" = $ResponseData.textvalue
                }
            }
            
            if ($ResponseData.ratingScore) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "score" = $ResponseData.ratingScore
                }
            }
            
            if ($ResponseData.additionalComment) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "additional.comment" = $ResponseData.additionalComment                
                }
                Write-Host $ResponseData.additionalComment -ForegroundColor DarkCyan
            }
            if ($($($ResponseData.selectOptions.label.en).Count -eq 1)) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "Option1" = "$($($ResponseData.selectOptions.label.en))" -replace '^"|"$'
                }
            }
            if ($($($ResponseData.selectOptions.label.en).Count -eq 2)) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "Option1" = "$($($ResponseData.selectOptions.label.en)[0])" -replace '^"|"$'
                    "Option2" = "$($($ResponseData.selectOptions.label.en)[1])" -replace '^"|"$'
                }
            }
            if ($($($ResponseData.selectOptions.label.en).Count -eq 3)) {
                $ThisObject."responseData"."$($SurveyQuestion.code)" += @{
                    "Option1" = "$($($ResponseData.selectOptions.label.en)[0])" -replace '^"|"$'
                    "Option2" = "$($($ResponseData.selectOptions.label.en)[1])" -replace '^"|"$'
                    "Option3" = "$($($ResponseData.selectOptions.label.en)[2])" -replace '^"|"$'
                }
            }
        }

        #Calculate first sentiment score
        $S1Score = (@([int]$($ThisObject.responseData.'lawfirm.question.9f331d9b'.score), 
                [int]$($ThisObject.responseData.'lawfirm.146.pe'.score),
                [int]$($ThisObject.responseData.'lawfirm.156.ex'.score),
                [int]$($ThisObject.responseData.'lawfirm.162.ue'.score)) | Measure-Object -Average).Average
        
        #Sentiment translation
        if ($S1Score -ge 4) {
            $S1Sentiment = "Favorable"
        }
        elseif ($S1Score -ge 3 -and $S1Score -lt 4) {
            $S1Sentiment = "Neutral"
        }
        elseif ($S1Score -lt 3) {
            $S1Sentiment = "Not Favorable"
        }  
        #Calculate second sentiment score
        $S2Score = (@([int]$($ThisObject.responseData.'lawfirm.question.9f331d9b'.score),
                [int]$($ThisObject.responseData.'lawfirm.question.377fc978'.score)) | Measure-Object -Average).Average
        
        if ($S2Score -ge 4) {
            $S2Sentiment = "Favorable"
        }
        elseif ($S2Score -ge 3 -and $S2Score -lt 4) {
            $S2Sentiment = "Neutral"
        }
        elseif ($S2Score -lt 3) {
            $S2Sentiment = "Not Favorable"
        }
        #Calculate third sentiment score
        $S3Score = (@([int]$($ThisObject.responseData.'lawfirm.148.dg'.score),
                [int]$($ThisObject.responseData.'lawfirm.164.dy'.score),
                [int]$($ThisObject.responseData.'lawfirm.question.f44611b8'.score),
                [int]$($ThisObject.responseData.'lawfirm.question.b3fff642'.score),
                [int]$($ThisObject.responseData.'lawfirm.question.b069a100'.score)) | Measure-Object -Average).Average
        
        if ($S3Score -ge 4) {
            $S3Sentiment = "Favorable"
        }
        elseif ($S3Score -ge 3 -and $S3Score -lt 4) {
            $S3Sentiment = "Neutral"
        }
        elseif ($S3Score -lt 3) {
            $S3Sentiment = "Not Favorable"
        }
        #Calculate overall seniment score
        $OverallScore = (@($S1Score, $S2Score, $S3Score) | Measure-Object -Average).Average
        
        if ($OverallScore -ge 4) {
            $OverallSentiment = "Favorable"
        }
        elseif ($OverallScore -ge 3 -and $OverallScore -lt 4) {
            $OverallSentiment = "Neutral"
        }
        elseif ($OverallScore -lt 3) {
            $OverallSentiment = "Not Favorable"
        }

        $ThisObject += @{
            "s1.qual.score" = "$($S1Sentiment)"
            "s2.qual.score" = "$($S2Sentiment)"
            "s3.qual.score" = "$($S3Sentiment)"
            "ov.qual.score" = "$($OverallSentiment)"
        }
        
        if (-not $Employees.ContainsKey($EmployeeData.name)) {
            $Employees[$EmployeeData.name] = @()
        }
        $Employees[$EmployeeData.name] += $ThisObject
        
        #Creates PDF file
        $SaveFile = Create-PDF $thisObject
        if ($SaveFile -ne $null) {
            #Upload generated PDF to SharePoint
            $FileURL = Upload-PDFsToSharePoint -FileLocation $SaveFile -ParentOffice $ThisObject.parentoffice -SubmittedAt $ThisObject.'submission.date.formatted'
            #Analyze survey sentiment
            $Sentiment = Get-SentimentAnalysis $ThisObject -Keywords $Keywords
            #Logs survey and sentiment information
            Log-RecordAndSentiment -EmployeeId $EmployeeData.employeeIdentifier -SurveyId $SurveyResponse.id -EmployeeName $EmployeeData.name -JsonData $Sentiment -SubmittedAt $SurveyResponse.submittedAt.ToString("yyyy-M-dd") -SurveyURL $FileURL.URL
        }
        Write-Output "Processed response for $($EmployeeData.name)" 
    }
    else {
        $EmployeeData = $(Get-EmployeeData -EmployeeId $SurveyResponse.employeeId)
        Write-Output "Response already processed for $($EmployeeData.name)" 
    }
}

# Consolidated email sending logic
try {
    # First send monthly survey results email notification
    Write-Output "Sending monthly survey results email notification..."
    $EmailResult = Send-GraphEmail -From $fromEmail -To $recipientEmails -Force:$Force -bcc $bccRecipients
    $EmailResult
    # Output results
    if ($EmailResult.Success) {
        Write-Output "Successfully sent $($EmailResult.EmailsSent) email(s)"
        Write-Output "Target folder: $($EmailResult.TargetFolder)"
    }
    else {
        Write-Output "Monthly email not sent: $($EmailResult.Reason)"
    }

    # Then send weekly sentiment review email to HR team
    Write-Output "Sending weekly sentiment review email to HR team..."
    $WeeklySentimentResult = Send-WeeklySentimentEmail -From $fromEmail -To $hrTeamEmails -Force:$Force
    
    # Output results
    if ($WeeklySentimentResult.Success) {
        Write-Output "Successfully sent weekly sentiment review email to $($WeeklySentimentResult.EmailsSent) HR recipient(s)"
        Write-Output "Week covered: $($WeeklySentimentResult.WeekStart) to $($WeeklySentimentResult.WeekEnd)"
        if ($WeeklySentimentResult.ResponseCount) {
            Write-Output "Responses analyzed: $($WeeklySentimentResult.ResponseCount)"
        }
    }
    else {
        Write-Output "Weekly sentiment email not sent: $($WeeklySentimentResult.Reason)"
    }

    # Return the monthly email result as before
    return $EmailResult
}
catch {
    Write-Error "Error in email sending operations: $_"
    throw $_
}

Disconnect-PnPOnline