param(
    [string]$EwsUrl = "https://mail.example.com/EWS/Exchange.asmx",
    [Parameter(Mandatory = $true)][string]$From,   # group address (you need Send As on it)
    [Parameter(Mandatory = $true)][string]$To
)

function Esc([string]$s) { [System.Security.SecurityElement]::Escape($s) }

$subject = "EWS test from $env:COMPUTERNAME"
$body    = "Test sent at $(Get-Date) from $env:COMPUTERNAME as $From over EWS (HTTPS), no SMTP and no Outlook."

$xml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:Subject>$(Esc $subject)</t:Subject>
          <t:Body BodyType="Text">$(Esc $body)</t:Body>
          <t:ToRecipients><t:Mailbox><t:EmailAddress>$(Esc $To)</t:EmailAddress></t:Mailbox></t:ToRecipients>
          <t:From><t:Mailbox><t:EmailAddress>$(Esc $From)</t:EmailAddress></t:Mailbox></t:From>
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
"@

try {
    $r = Invoke-WebRequest -Uri $EwsUrl -Method Post -ContentType "text/xml; charset=utf-8" -Body $xml -UseDefaultCredentials -UseBasicParsing
    if ($r.Content -match 'ResponseClass="Success"') { Write-Host "SENT OK via EWS: $From -> $To" }
    else { Write-Warning "EWS answered but did not report success:"; $r.Content }
}
catch {
    Write-Warning "EWS call failed: $($_.Exception.Message)"
    if ($_.Exception.Response) { Write-Warning ("HTTP status: " + [int]$_.Exception.Response.StatusCode) }
}
