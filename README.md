# PowerShell-Podcast

A quick-and-dirty podcast downloader written in PowerShell.  In lieu of a user-interface, it is controlled by CSV files that you can edit in your spreadsheet software of choice.

It was created to meet my needs: a podcast downloader that downloads audio files with human-readable filenames into a directory that I can sync to my digital audio player.  I normally listen on a tiny [Sansa c200](http://en.wikipedia.org/wiki/Sansa_c200_series) and not a smartphone, an iPod, or a Zune, so I want a tool that treats my music player like a thumb drive.  I use a syncing tool like [Synkron](http://synkron.sourceforge.net/) to copy the podcasts to my digital audio player, and I use [OneDrive](https://onedrive.live.com/about/en-us/) to sync between computers and onto a smartphone.

**Note that this script is still incomplete but it works well enough for me.**

## Usage:

Place the main script in an empty directory.  It will store configuration and downloaded podcasts in this directory.

Run it from the command-line:
```
> .\powershell-podcast
```

It will create an empty `subscriptions.csv` file.  Open this file in your favorite spreadsheet software.

ID | Name | Directory | URL | DownloadQuantity
---|------|-----------|-----|-----------------
 | | | | 
 | | | | 
 | | | | 

Add a few rows for the podcasts you want to subscribe to, and save the file.

* Name and ID can be anything you want.
* Directory is the name of the subdirectory into which this podcast will be downloaded.
* DownloadQuantity is the number of most recent podcasts that will be downloaded from this feed.  For example, if this is 1, only the single most recent podcast will be downloaded every time this script is run.

Run the script again:
```
.\powershell-podcast
```

It will fetch your podcasts' feeds and start downloading audio files.

To periodically download more podcasts, periodically run the script again.  A separate CSV file is created for each subscription to track which files have been downloaded.  This means that the script will never download an audio file more than once.  You can safely delete old files once you're done listening to them.
