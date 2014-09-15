$ErrorActionPreference = "Stop"

function Main {

    # CD into the directory where this script resides
    Push-Location $PSScriptRoot

    # Create an empty subscription list if it doesn't exist
    if ((Test-Path './subscriptions.csv') -eq $false) {
        Write-Output "ID,Name,Directory,URL,DownloadQuantity" > './subscriptions.csv'
    }

    # Read the list of subscriptions
    $subscriptions = Get-Content ./subscriptions.csv | ConvertFrom-Csv

    foreach($subscription in $subscriptions) {
        
        Write-Output "Fetching feed for $($subscription.Name)..."
        
        # Download the feed and parse it as XML
        [xml]$feed = Invoke-WebRequest $subscription.URL

        # Load the list of existing entries for this podcast from the filesystem
        $existingEntriesPath = "./podcasts/$($subscription.directory).csv"
        # Create an empty list if it doesn't exist
        if ((Test-Path $existingEntriesPath) -eq $false) {
            Write-Output "GUID,Title,Date,URL,Filename,Downloaded,ListenedTo" > $existingEntriesPath
        }
        $existingEntries = Get-Content $existingEntriesPath | ConvertFrom-Csv

        # Parse a list of potentially new entries from the downloaded feed's XML
        $feedEntries = createPodcastEntriesFromFeedXml $feed

        # Merge all RSS feed entries into the list of existing entries
        $newEntries = foreach($feedEntry in $feedEntries) {
            if($existingEntries | where GUID -eq $feedEntry.GUID) {
                # Entry already exists; do nothing
            } else {
                Write-Output $feedEntry
            }
        }

        # A strange way to concatenate in PowerShell
        # Also, remove null values
        $allEntries = Write-Output $newEntries $existingEntries | where { $_ }
        
        # Sort all entries by date, with newest entries at the top
        $allEntries = $allEntries | Sort-Object -Property Date -Descending

        # write the CSV to disc
        writeEntriesToDisc $allEntries $existingEntriesPath

        # Based on user preferences for this podcast, figure out which new entries must be downloaded
        $maxDownloaded = [int]$subscription.DownloadQuantity

        # TODO change this algorithm to take two numbers into account:
        # How many of the most recent podcasts to have downloaded,
        # and the absolute maximum number to have downloaded.
        # For example, if the absolute max is 6, you already have 5 downloaded, and 3 more new podcasts have appeared in the feed,
        # we should only download one more.  Only after older ones have been deleted can we download more without exceeding our absolute max limit.

        if($maxDownloaded -le 0) {
            $entriesToDownload = @()
        } else {
            $entriesToDownload = ($allEntries | where { isNo $_.ListenedTo })[0..($maxDownloaded - 1)] | where { isNo $_.Downloaded }
            # Coerce it to an array, in case a single value was returned
            $entriesToDownload = @($entriesToDownload)
        }

        # Make the podcast output directory if it doesn't exist
        $directory = "./podcasts/$($subscription.Directory)"
        if((Test-Path $directory) -eq $false) {
            mkdir $directory
        }

        if ($entriesToDownload.Count -gt 0) {
            Write-Output "Downloading $($entriesToDownload.Count) podcast$(if($entriesToDownload.Count -gt 1) {"s"}) for $($subscription.Name)..."
        }

        # Download each podcast into this podcast's directory and mark it as downloaded in the CSV
        foreach($entryToDownload in $entriesToDownload) {
            $filename = "$directory/$($entryToDownload.Filename)"
            Invoke-WebRequest $entryToDownload.URL -OutFile $filename
            $entryToDownload.Downloaded = "Yes"
            
            # Write the CSV to disc
            writeEntriesToDisc $allEntries $existingEntriesPath
        }
        
        # TODO
        # Based on the CSV, delete old entries?
        # Based on the filesystem, delete old entries?
        #   If downloaded == true and it's not on the filesystem, assume I listened to it

        # Write the CSV to disc
        writeEntriesToDisc $allEntries $existingEntriesPath
    }
    
    Pop-Location
}

# Write a list of podcast entries to disc.  We do this fairly often so that, no matter when the script is killed, it has output
# an accurate representation of state to the hard drive.
function writeEntriesToDisc($allEntries, $path) {
    $allEntries | ConvertTo-Csv -NoTypeInformation > $path
}

function createPodcastEntriesFromFeedXml($xml) {
    # Create and return an array of PodcastEntry from all the items in this XML dom
    $entryNodes = Select-Xml -XPath "//item" -Xml $feed
    $entryNodes | % {
        $entryNode = $_.Node
        $entry = createPodcastEntry
        $entry.Title = (Select-Xml -XPath "title/text()" -Xml $entryNode).Node.InnerText
        
        # Parse the publication date and format it as a UTC string
        $date = (Select-Xml -XPath "pubDate/text()" -Xml $entryNode).Node.InnerText
        $entry.Date = (parseDate $date).ToUniversalTime().ToString('yyyy-MM-ddThh:mm:ss.fffK')
        
        $entry.URL = (Select-Xml -XPath "enclosure/@url" -Xml $entryNode).Node.Value

        # Use the <item>'s <guid> element.  If it's not present, fall back to using the download URL as a GUID.
        $guid = (Select-Xml -XPath "guid/text()" -Xml $entryNode).Node.InnerText
        if($guid -eq $null) {
            $guid = $entry.URL
        }
        $entry.GUID = $guid
        
        # TODO make this more robust?
        # If the mimetype is not audio/mpeg, do not make it .mp3?
        # If the URL has a file extension at the end, use that?
        $entry.filename = "$(sanitizeFilename $entry.title).mp3"
        return $entry
    }
}

function createPodcastEntry {
    [pscustomobject]@{
        GUID=""
        Title=""
        Date=""
        URL=""
        Filename=""
        Downloaded=$false
        ListenedTo=$false
    }
}

function isNo($yesOrNoString) {
    $yesOrNoString -Match "^[^Yy]"
}

function isYes($yesOrNoString) {
    $yesOrNoString -Match "^[Yy]"
}

$invalidFileNameChars = [System.IO.Path]::GetInvalidFileNameChars()
function sanitizeFilename($name) {
    $ret = $name
    foreach($char in $invalidFileNameChars) {
        # TODO perform url-encode-style percent-encoding
        # TODO use linq aggregate function to do all this find-and-replace
        $ret = $ret -Replace [regex]::Escape($char),"_"
    }
    return $ret
}

# Parses an RFC822 string into a DateTime object
function parseDate($date) {
    $timezones = @{
        "UT" = "GMT"
        "GMT" = "GMT"
        "EST" = "-0500"
        "EDT" = "-0400"
        "CST" = "-0600"
        "CDT" = "-0500"
        "MST" = "-0700"
        "MDT" = "-0600"
        "PST" = "-0800"
        "PDT" = "-0700"
        "Z" = "GMT"
        "A" = "-0100"
        "M" = "-1200"
        "N" = "+0100"
        "Y" = "+1200"
    }
    $acc = $date
    foreach($abbrev in $timezones.Keys) {
        $acc = $acc -replace "$abbrev$",$timezones[$abbrev]
    }
    [DateTime]$acc
}

return Main
