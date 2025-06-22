<#
.SYNOPSIS
    Connects to the Google Gemini API to stream generative content.

.DESCRIPTION
    This script facilitates interaction with the Google Gemini API to generate
    text content based on a provided prompt. It supports streaming responses
    and includes error handling for API-related issues.

.PARAMETER ApiKey
    Your Google Gemini API Key. This is required for authentication with the API.

.PARAMETER Model
    The specific Gemini model to use for content generation (e.g., 'gemini-1.5-flash').

.PARAMETER Prompt
    The text prompt to send to the Gemini API for content generation.

.PARAMETER DebugMode
    Enables debug output for troubleshooting.

.NOTES
    Version: 1.0
    Author: PowerShell Developer
    Date: 2023-10-27
    Requires: PowerShell 5.1 or higher
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$ApiKey,

    [Parameter(Mandatory=$false)]
    [string]$Model = "gemini-1.5-flash", # Default model

    [Parameter(Mandatory=$false)]
    [string]$Prompt,

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

# --- Script Version ---
$ScriptVersion = "1.0"

# --- Function for Colored Output ---
function Write-ColoredHost {
    param (
        [string]$Message,
        [ConsoleColor]$Color
    )
    Write-Host $Message -ForegroundColor $Color
}

# --- Input Validation and Read-Host Prompts ---
if (-not $ApiKey) {
    Write-ColoredHost "Please enter your Google Gemini API Key:" White
    $ApiKey = Read-Host
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        Write-ColoredHost "API Key cannot be empty. Exiting script." Red
        exit 1
    }
}

if (-not $Prompt) {
    Write-ColoredHost "Please enter the prompt for Gemini (e.g., 'Write a short story about a robot who discovers music.'):" White
    $Prompt = Read-Host
    if ([string]::IsNullOrWhiteSpace($Prompt)) {
        Write-ColoredHost "Prompt cannot be empty. Exiting script." Red
        exit 1
    }
}

# --- API Request Details ---
# API Key is now passed as a query parameter 'key'
$apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/$($Model):streamGenerateContent?key=$ApiKey&alt=sse"

$body = @{
    contents = @(
        @{
            parts = @(
                @{
                    text = $Prompt
                }
            )
        }
    )
} | ConvertTo-Json -Depth 4

# --- Debugging Output ---
if ($DebugMode) {
    Write-ColoredHost "--- Debug Information ---" Cyan
    Write-ColoredHost "API URL: $apiUrl" Cyan
    Write-ColoredHost "Request Body:" Cyan
    Write-ColoredHost $body Cyan
    Write-ColoredHost "-------------------------" Cyan
}

# --- Make the Streaming API Call ---
$streamReader = $null
$responseStream = $null

try {
    Write-ColoredHost "Sending request to Gemini API..." Cyan
    $webRequest = [System.Net.WebRequest]::Create($apiUrl)
    $webRequest.Method = "POST"
    $webRequest.ContentType = "application/json"

    $requestBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $webRequest.ContentLength = $requestBytes.Length

    $requestStream = $webRequest.GetRequestStream()
    $requestStream.Write($requestBytes, 0, $requestBytes.Length)
    $requestStream.Close()

    # --- Read the Streaming Response ---
    $response = $webRequest.GetResponse()
    $responseStream = $response.GetResponseStream()
    $streamReader = New-Object System.IO.StreamReader($responseStream)

    Write-ColoredHost "Gemini is thinking..." Yellow
    while (-not $streamReader.EndOfStream) {
        $line = $streamReader.ReadLine()
        if ($line.StartsWith("data: ")) {
            $jsonData = $line.Substring(6)
            try {
                $data = $jsonData | ConvertFrom-Json
                # Check if text exists before trying to access it
                if ($data.candidates[0].content.parts[0].text) {
                    $text = $data.candidates[0].content.parts[0].text
                    Write-Host $text -NoNewline # -NoNewline for continuous output
                }
            }
            catch {
                # This can happen on the last, empty data chunk or malformed JSON.
                # If DebugMode is on, show the error.
                if ($DebugMode) {
                    Write-ColoredHost "Error parsing JSON data stream: $_" Yellow
                    Write-ColoredHost "Problematic JSON: $jsonData" Yellow
                }
            }
        }
    }
    # Add a newline at the end for clean formatting
    Write-Host ""
    Write-ColoredHost "Gemini response complete." Green
}
catch {
    # Attempt to read the error response from the server
    if ($_.Exception.Response) {
        $errorResponseStream = $_.Exception.Response.GetResponseStream()
        $errorStreamReader = New-Object System.IO.StreamReader($errorResponseStream)
        $errorBody = $errorStreamReader.ReadToEnd()
        Write-ColoredHost "An API error occurred. Response: $errorBody" Red
        $errorStreamReader.Close()
    } else {
        Write-ColoredHost "An unexpected error occurred: $($_.Exception.Message)" Red
        if ($DebugMode) {
            Write-ColoredHost "Full error details:" Red
            Write-Error $_ -ErrorAction SilentlyContinue
        }
    }
}
finally {
    # Ensure streams are closed
    if ($streamReader) {
        $streamReader.Close()
        if ($DebugMode) { Write-ColoredHost "StreamReader closed." Cyan }
    }
    if ($responseStream) {
        $responseStream.Close()
        if ($DebugMode) { Write-ColoredHost "ResponseStream closed." Cyan }
    }
    Write-ColoredHost "Script execution finished." Cyan
}