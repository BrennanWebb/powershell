# --- Configuration ---
$apiKey = "APIKey"
$model = "gemini-1.5-flash" # Or any other supported model
$prompt = "Write a short story about a robot who discovers music."

# --- API Request Details ---
$apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/$($model):streamGenerateContent?alt=sse"

$body = @{
    contents = @(
        @{
            parts = @(
                @{
                    text = $prompt
                }
            )
        }
    )
} | ConvertTo-Json -Depth 4

# --- Make the Streaming API Call ---
try {
    $webRequest = [System.Net.WebRequest]::Create($apiUrl)
    $webRequest.Method = "POST"
    
    # --- FIX START ---
    # Set headers individually instead of adding a hashtable object.
    $webRequest.Headers.Add("Authorization", "Bearer $apiKey")
    $webRequest.ContentType = "application/json"
    # --- FIX END ---

    $requestBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $webRequest.ContentLength = $requestBytes.Length

    $requestStream = $webRequest.GetRequestStream()
    $requestStream.Write($requestBytes, 0, $requestBytes.Length)
    $requestStream.Close()

    # --- Read the Streaming Response ---
    $response = $webRequest.GetResponse()
    $responseStream = $response.GetResponseStream()
    $streamReader = New-Object System.IO.StreamReader($responseStream)

    Write-Host "Gemini is thinking..." -ForegroundColor Yellow
    while (-not $streamReader.EndOfStream) {
        $line = $streamReader.ReadLine()
        if ($line.StartsWith("data: ")) {
            $jsonData = $line.Substring(6)
            try {
                $data = $jsonData | ConvertFrom-Json
                # Check if text exists before trying to access it
                if ($data.candidates[0].content.parts[0].text) {
                    $text = $data.candidates[0].content.parts[0].text
                    Write-Host $text -NoNewline
                }
            }
            catch {
                # This can happen on the last, empty data chunk. Silently continue.
            }
        }
    }
    # Add a newline at the end for clean formatting
    Write-Host ""
}
catch {
    # Attempt to read the error response from the server
    if ($_.Exception.Response) {
        $errorResponseStream = $_.Exception.Response.GetResponseStream()
        $errorStreamReader = New-Object System.IO.StreamReader($errorResponseStream)
        $errorBody = $errorStreamReader.ReadToEnd()
        Write-Error "An API error occurred. Response: $errorBody"
        $errorStreamReader.Close()
    } else {
         Write-Error "An error occurred: $_"
    }
}
finally {
    if ($streamReader) { $streamReader.Close() }
    if ($responseStream) { $responseStream.Close() }
}