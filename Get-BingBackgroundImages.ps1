<#
Get-BingBackgroundImages.ps1
#>


[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
      HelpMessage="`r`nOutput: In which folder or directory would you like to find the new Bing backgrounds? `r`n`r`nPlease enter a valid file system path to a directory (a full path name of a folder such as C:\Windows). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path in quotation marks (single or double). `r`n`t- To exit this script, please press [Ctrl] + C `r`n")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [Alias("Path","Output","OutputFolder")]
    [string]$Destination = "$($env:USERPROFILE)\Pictures\Wallpapers\Bing",
    [ValidateSet("en-US","zh-CN","ja-JP","en-AU","en-UK","de-DE","en-NZ","en-CA")]
    [string]$Market = "en-US",
    [ValidateRange(0,15)]
    [int]$Index = 0,
    [ValidateRange(1,8)]
    [Alias("Number","Amount","Images")]
    [int]$NumberOfImages = 1,
    [string]$Resolution = "1920x1080",
    [Alias("SubfolderForThePortraitPictures","SubfolderForTheVerticalPictures","SubfolderName")]
    [string]$Subfolder = "Vertical",
    [switch]$Log,
    [Alias("Include")]
    [switch]$IncludePortrait
)


Begin {

    # Function used to convert bytes to MB or GB or TB                                        # Credit: clayman2: "Disk Space"
    function ConvertBytes {
        Param (
            $size
        )
        If ($size -eq $null) {
            [string]'-'
        } ElseIf ($size -eq 0) {
            [string]'-'
        } ElseIf ($size -lt 1MB) {
            $file_size = $size / 1KB
            $file_size = [Math]::Round($file_size, 0)
            [string]$file_size + ' KB'
        } ElseIf ($size -lt 1GB) {
            $file_size = $size / 1MB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' MB'
        } ElseIf ($size -lt 1TB) {
            $file_size = $size / 1GB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' GB'
        } Else {
            $file_size = $size / 1TB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' TB'
        } # else
    } # function (ConvertBytes)


    # Set the common parameters                                                               # Credit: Jeffrey: "Get-BingImage.ps1"
    # Source: http://stackoverflow.com/questions/10639914/is-there-a-way-to-get-bings-photo-of-the-day
    # Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string
    $language       = ($Market.Split("-")[0]).ToLower()
    $region         = ($Market.Split("-")[-1]).ToUpper()
    $real_market    = [string]$language + '-' + $region
    $bing_url       = "http://www.bing.com"
    $query          = "&idx=$($Index)&n=$($NumberOfImages)&mkt=$($real_market)"
    $data_url       = "$($bing_url)/HPImageArchive.aspx?format=xml$($query)"
    $xml_file       = "$env:temp\bing.xml"
    $existing_files = 0
    $downloaded     = 0
    $images         = @()
    $empty_line     = ""
    $date           = Get-Date -Format g
    $timestamp      = Get-Date -Format yyyyMMdd
    $target_images  = If ($IncludePortrait)         { [int]($NumberOfImages * 2) }                      Else { [int]($NumberOfImages) }
    If ($target_images -eq 1)                       { $word = "image" }                                 Else { $word = "images" }
    If ((($Destination).EndsWith("\")) -eq $true)   { $Destination = $Destination -replace ".{1}$" }    Else { $continue = $true }



    # Check if the computer is connected to the Internet                                      # Credit: ps1: "Test Internet connection"
    If (([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet) -eq $false) {
        $empty_line | Out-String
        Return "The Internet connection doesn't seem to be working. Exiting without downloading any Bing background images."
    } Else {
        $continue = $true
    } # else


    try
    {
        # Set the progress bar variables ($id denominates different progress bars, if more than one is being displayed)
        $activity           = "Downloading $target_images Bing background images"
        $status             = " "
        $task               = "Setting Initial Variables"
        $total_steps        = [int]($target_images + 2)
        $activities         = [int]($target_images + 1)
        $task_number        = 0.2
        $image_num          = 0
        $step               = 1
        $id                 = 1


        # Start the progress bar
        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)

            # Determine the original picture orientation
            $regex = $Resolution -match "(?<R1>\d+)x(?<R2>\d+)"
            $width      = If ($Matches.R1 -ne $null) { [int]$Matches.R1 } Else { Return "'$Resolution' is either not a valid resolution or it is not recommended." }
            $height     = If ($Matches.R2 -ne $null) { [int]$Matches.R2 } Else { Return "'$Resolution' is either not a valid resolution or it is not recommended." }

            If ($width -gt $height) {
                $type = "Landscape"
                $orientation = "Horizontal"
                $destination_path = "$Destination"
                $subfolder_path = "$Destination\$Subfolder"
                $portrait_resolution = [string]$height + 'x' + $width

                    If ($IncludePortrait) {
                        If ((Test-Path $subfolder_path -PathType 'Container') -eq $true) {
                            $continue = $true
                        } Else {
                            New-Item "$subfolder_path" -ItemType Directory -Force | Out-Null
                        } # Else (If Test-Path $subfolder_path)
                    } Else {
                        $continue = $true
                    } # Else (If $IncludePortrait)

            } ElseIf ($height -gt $width) {
                $type = "Portrait"
                $orientation = "Vertical"
                $destination_path = "$Destination\$Subfolder"

                    # Create a subfolder for the portrait (vertical) pictures, if it doesn't exist
                    If ((Test-Path $destination_path -PathType 'Container') -eq $true) {
                        $continue = $true
                    } Else {
                        New-Item "$destination_path" -ItemType Directory -Force | Out-Null
                    } # Else (If Test-Path $destination_path)
            } Else {
                $continue = $true
            } # Else (If $width)
    }
    catch
    {
        Write-Error $Error[0].Exception
        Break
    }
} # Begin




Process {

        # Update the progress bar
        $activity = "Downloading the Bing XML-file - Step $step/$activities"
        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation "Source: $data_url" -PercentComplete (($step / $total_steps) * 100)

    try
    {
        # Download the Bing XML-file containing the image info                                # Credit: Jeffrey: "Get-BingImage.ps1"
        # Source: Tobias Weltner: "PowerTips Monthly vol 10 March 2014" (Read raw web page content)
        $image_crawler = New-Object System.Net.WebClient
        $web_client = New-Object System.Net.WebClient
        $raw_bing_data = $web_client.DownloadString($data_url)
        $raw_bing_data | Out-File "$xml_file" -Encoding Default
        [xml]$bing_info = Get-Content $xml_file -Encoding UTF8


        ForEach ($image in $bing_info.images.image) {

            # Source: https://msdn.microsoft.com/en-us/library/system.uri(v=vs.110).aspx
            [System.Uri]$download_uri = New-Object System.Uri "$($bing_url)$($image.urlBase)_$($Resolution).jpg"
            $download_url = $download_uri.AbsoluteUri
            $original_filename = $download_uri.Segments[-1]
            $filename = "$($image.startdate)_$($real_market)_$($original_filename.Split('_')[0])_$($original_filename.Split('_')[-1])"
            $root_name = ($original_filename.Split('_')[0])
            $filepath = "$($destination_path)\$($filename)"
            $title = (($image.copyright.Split("©")[0]).TrimEnd(" ("))
            $author = ((($image.copyright.Split("©")[-1]).Split("/")[0]).TrimStart(" "))
            $copyright = ((($image.copyright.Split("©")[-1]).Replace(')','')).TrimStart(" "))
            $custom_url = [string]"http://www.bing.com/gallery/#images/" + $root_name
            $copyright_link = If (($image.copyrightlink -match 'javascript:void') -eq $true) { "-" } Else { $image.copyrightlink }
            If (($image.copyrightlink -match "HpDate:") -eq $true) {
                $real_date = (($image.copyrightlink.Replace('%22','')).Split("HpDate:")[-1])
            } Else {
                $real_date = $image.startdate
            } # Else (If $image.copyrightlink)

            # Increment the counters
            $image_num++
            $step++

            # Update the progress bar
            $activity = "Downloading $target_images Bing background $word - Step $step/$activities"
            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $original_filename -PercentComplete (($step / $total_steps) * 100)

                If ((Test-Path $filepath -PathType 'Leaf') -eq $true) {
                    Write-Verbose "Skipping $filepath, which seems to already exist."
                    $existing_files++
                } Else {

                    # Download the Bing background image
                    # Source: Tobias Weltner: "PowerTips Monthly vol 10 March 2014" (Read raw web page content)
                    $image_data = $image_crawler.DownloadString($download_url)
                    $image_data | Out-File "$filepath" -Encoding Default
                    $file = Get-ChildItem -Path "$filepath" -Force -ErrorAction SilentlyContinue
                    $downloaded++

                    $images += $obj_image = New-Object -TypeName PSCustomObject -Property @{
                            'ID'                        = $image_num
                            'File'                      = $filename
                            'Date'                      = $real_date
                            'End Date'                  = $image.enddate
                            'end_date'                  = $image.enddate
                            'Download Date'             = $timestamp
                            'Resolution'                = $Resolution
                            'Type'                      = $type
                            'Size'                      = (ConvertBytes ($file.Length))
                            'Title'                     = $title
                            'Author'                    = $author
                            'Photo Agency'              = $copyright.Replace("$author/","")
                            'Copyright'                 = (($image.copyright.Replace("$title (","")).TrimEnd(")"))
                            'Copyright (Full)'          = $image.copyright
                            'Custom URL'                = $custom_url
                            'Copyright Link'            = $copyright_link
                            'drk'                       = $image.drk
                            'top'                       = $image.top
                            'bot'                       = $image.bot
                            'hotspots'                  = $image.hotspots
                            'Orientation'               = $orientation
                            'raw_size'                  = $file.Length
                            'Host'                      = $download_uri.Host
                            'URL (base)'                = $image.urlBase
                            'URL (raw)'                 = $image.url
                            'Original URL'              = $download_uri.OriginalString
                            'Source'                    = $download_url
                            'Destination Folder'        = $destination_path
                            'Destination'               = $filepath
                            'Root Name'                 = $root_name
                            'Identifier Code'           = ($image.urlBase.Split('_')[-1])
                            'Market'                    = $real_market
                            'Authority'                 = $download_uri.Authority
                            'Width'                     = $width
                            'Height'                    = $height
                            'Downloaded'                = $date
                            'Start Date'                = $image.startdate
                            'Full Start Date'           = $image.fullstartdate
                        } # New-Object
                } # Else (If Test-Path $filepath)


                # Process also portrait pictures, if set to do so with the -IncludePortrait parameter
                If (($type -eq "Landscape") -and ($IncludePortrait)) {

                    # Source: https://msdn.microsoft.com/en-us/library/system.uri(v=vs.110).aspx
                    [System.Uri]$portrait_download_uri = New-Object System.Uri "$($bing_url)$($image.urlBase)_$portrait_resolution.jpg"
                    $portrait_download_url = $portrait_download_uri.AbsoluteUri
                    $original_portrait_filename = $portrait_download_uri.Segments[-1]
                    $portrait_filename = "$($image.startdate)_$($real_market)_$($original_portrait_filename.Split('_')[0])_$($original_portrait_filename.Split('_')[-1])"
                    $portrait_filepath = "$($subfolder_path)\$($portrait_filename)"

                    # Increment the counters
                    $image_num++
                    $step++

                    # Update the progress bar
                    $activity = "Downloading $target_images Bing background $word - Step $step/$activities"
                    Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $original_portrait_filename -PercentComplete (($step / $total_steps) * 100)

                        If ((Test-Path $portrait_filepath -PathType 'Leaf') -eq $true) {
                            Write-Verbose "Skipping $portrait_filepath, which seems to already exist."
                            $existing_files++
                        } Else {

                            # Download the portrait version of the Bing background image
                            # Source: Tobias Weltner: "PowerTips Monthly vol 10 March 2014" (Read raw web page content)
                            $portrait_data = $image_crawler.DownloadString($portrait_download_url)
                            $portrait_data | Out-File "$portrait_filepath" -Encoding Default
                            $portrait_file = Get-ChildItem -Path "$portrait_filepath" -Force -ErrorAction SilentlyContinue
                            $downloaded++

                            $images += $obj_portrait = New-Object -TypeName PSCustomObject -Property @{
                                    'ID'                        = $image_num
                                    'File'                      = $portrait_filename
                                    'Date'                      = $real_date
                                    'End Date'                  = $image.enddate
                                    'end_date'                  = $image.enddate
                                    'Download Date'             = $timestamp
                                    'Resolution'                = $portrait_resolution
                                    'Type'                      = "Portrait"
                                    'Size'                      = (ConvertBytes ($portrait_file.Length))
                                    'Title'                     = $title
                                    'Author'                    = $author
                                    'Photo Agency'              = $copyright.Replace("$author/","")
                                    'Copyright'                 = (($image.copyright.Replace("$title (","")).TrimEnd(")"))
                                    'Copyright (Full)'          = $image.copyright
                                    'Custom URL'                = $custom_url
                                    'Copyright Link'            = $copyright_link
                                    'drk'                       = $image.drk
                                    'top'                       = $image.top
                                    'bot'                       = $image.bot
                                    'hotspots'                  = $image.hotspots
                                    'Orientation'               = "Vertical"
                                    'raw_size'                  = $portrait_file.Length
                                    'Host'                      = $portrait_download_uri.Host
                                    'URL (base)'                = $image.urlBase
                                    'URL (raw)'                 = $image.url
                                    'Original URL'              = $portrait_download_uri.OriginalString
                                    'Source'                    = $portrait_download_url
                                    'Destination Folder'        = $subfolder_path
                                    'Destination'               = $portrait_filepath
                                    'Root Name'                 = $root_name
                                    'Identifier Code'           = ($image.urlBase.Split('_')[-1])
                                    'Market'                    = $real_market
                                    'Authority'                 = $portrait_download_uri.Authority
                                    'Width'                     = $height
                                    'Height'                    = $width
                                    'Downloaded'                = $date
                                    'Start Date'                = $image.startdate
                                    'Full Start Date'           = $image.fullstartdate
                                } # New-Object
                        } # Else (If Test-Path $portrait_filepath)
                } Else {
                    $continue = $true
                } # Else (If $IncludePortrait)
        } # ForEach
    }
    catch
    {
        Write-Error $Error[0].Exception
        Break
    }
} # Process




End {
                        # Close the progress bar
                        $task = "Finished downloading Bing background images."
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($total_steps / $total_steps) * 100) -Completed

    If ($images.Count -ge 1) {

        # Display the results in a pop-up window (Out-GridView)
        $images.PSObject.TypeNames.Insert(0,"Bing Background Images")
        $images_selection = $images | Sort end_date,Width -Descending | Select-Object 'ID','File','Date','End Date','Download Date','Resolution','Type','Size','Title','Author','Photo Agency','Copyright','Copyright (Full)','Custom URL','Copyright Link','drk','top','bot','hotspots','Orientation','raw_size','Host','URL (base)','URL (raw)','Original URL','Source','Destination Folder','Destination','Root Name','Identifier Code','Market','Authority','Width','Height','Downloaded','Start Date','Full Start Date'
        $images_selection | Out-GridView

                # Make the log entry if set to do so with the -Log parameter
                # Note: Append parameter of Export-Csv was introduced in PowerShell 3.0.
                # Source: http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
                # Source: https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/
                If ($Log) {
                    $logfile_path = "$Destination\bing_log.csv"
                    If ((Test-Path $logfile_path) -eq $false) {
                        $images_selection | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8
                    } Else {
                        # $images_selection | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Append
                        $images_selection | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $logfile_path -Append -Encoding UTF8
                        (Get-Content $logfile_path) | ForEach-Object { $_ -replace ('"', '') } | Out-File -FilePath $logfile_path -Force -Encoding UTF8
                    } # Else (If Test-Path $logfile_path)
                } Else {
                    $continue = $true
                } # Else (If $Log)
    } Else {
        $continue = $true
    } # Else (If $images.Count)


                        # Display rudimentary stats in console
                        If (($downloaded -ge 2) -or ($downloaded -eq 0)) {
                            $download_word = "images"
                        } ElseIf ($downloaded -eq 1) {
                            $download_word = "image"
                        } Else {
                            $continue = $true
                        } # Else (If $downloaded)

                        If (($existing_files -ge 2) -or ($existing_files -eq 0)) {
                            $existing_word = "images"
                        } ElseIf ($existing_files -eq 1) {
                            $existing_word = "image"
                        } Else {
                            $continue = $true
                        } # Else (If $existing_files)
                        $empty_line | Out-String
                        $text = "Downloaded $downloaded new $download_word. Didn't touch the $existing_files existing $existing_word."
                        Write-Output $text
                        $empty_line | Out-String
} # End




# [End of Line]


<#

   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27                                                 # Jeffrey: "Get-BingImage.ps1"
http://powershell.com/cs/media/p/7476.aspx                                                                  # clayman2: "Disk Space"
http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx                        # ps1: "Test Internet connection"
http://powershell.com/cs/media/p/32274.aspx                                                                 # Tobias Weltner: "PowerTips Monthly vol 10 March 2014" (Read raw web page content)


  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#
.SYNOPSIS
Downloads Bing background images in a defined resolution.

.DESCRIPTION
Get-BingBackgroundImages downloads a Bing XML-file (bing.xml) from Bing
http://www.bing.com/HPImageArchive.aspx?format=xml&idx=[Index]&n=[NumberOfImages]&mkt=[Market]
http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=en-US containing
info about the queried daily background images, which were recently featured on
the Bing search engine homepage (http://www.bing.com/) as a daily changing
background image. Microsoft's Bing search engine is famous for using hand-picked
beautiful imagery on their main landing page. The daily background images regularly
depict stunning landmarks, inspiring nature shots and high quality pictures of
people and cultures from around the world.

After determining the save locations, which are set with -Destination and -Subfolder
parameters, Get-BingBackgroundImages compares the filenames found in the destination
folders to the filenames listed in the XML-file (bing.xml) and tries to determine,
whether downloading an image is actually needed or not. Get-BingBackgroundImages
tries to avoid any unnecessary downloading and overwriting of the existing files.

By using the Bing API with AJAX calls it seems to be possible to retrieve up to 14
days old pictures and up to 8 images at a time. To retrieve all currently available
Bing background images in the default 1920x1080 resolution along with the 1080x1920
portrait variants as well from the default Bing market (en-US) with
Get-BingBackgroundImages and to create/update a log file (bing_log.csv), please
use two separate commands launching Get-BingBackgroundImages:

    .\Get-BingBackgroundImages.ps1 -Index 0 -NumberOfImages 7 -Log -IncludePortrait
    .\Get-BingBackgroundImages.ps1 -Index 7 -NumberOfImages 8 -Log -IncludePortrait

Please note that Get-BingBackgroundImages proceeds towards the history, when handling
multiple images. Please also note that if any of the individual parameter values
include space characters, the individual value should be enclosed in quotation marks
(single or double) so that PowerShell can interpret the command correctly. This
script is based on Jeffrey's script "Get-BingImage.ps1"
(https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27).

.PARAMETER Destination
with aliases -Path, -Output and -OutputFolder. Specifies the primary folder,
where the acquired new Bing background images are to be saved, and defines the
default location to be used with the -Log parameter (bing_log.csv), and also sets
the parent directory for the -Subfolder parameter. The default save location for
the horizontal (landscape) Bing background images is
"$env:USERPROFILE\Pictures\Wallpapers\Bing", which will be used, if no value
for the -Destination parameter is defined in the command launching
Get-BingBackgroundImages. For best results in iterative usage of this script,
the default value should remain constant and be set according to the prevailing
conditions (at line 13). The value for the -Destination parameter should be a valid
file system path pointing to a directory (a full path of a folder such as
C:\Users\Dropbox\). Furthermore, if the path includes space characters, please
enclose the path in quotation marks (single or double).

.PARAMETER Market
Sets the Bing market, from where the images are downloaded. The valid values are:
en-US, zh-CN, ja-JP, en-AU, en-UK, de-DE, en-NZ and en-CA. If no value for the
-Market parameter is defined in the command launching Get-BingBackgroundImages,
the default value "en-US" will be used.

.PARAMETER Index
Sets the start date, from which point towards the history should the images be
retrieved. 0 denominates the current day, 1 denominates yesterday, 2 denominates
the day before yesterday etc. It should be noted that Get-BingBackgroundImages
always counts "backwads" (towards the history), when retrieving multiple Bing
background images, and that by default, Get-BingBackgroundImages starts to count
the images from the current day.

.PARAMETER NumberOfImages
with aliases -Number, -Amount and -Images. Sets the number of images to return.
n = 1 would return only one, n = 2 would return two, and so on up to eight. By
default Get-BingBackgroundImages retrieves one image. It should also be noted that
Get-BingBackgroundImages always counts "backwads" (towards the history), when
retrieving Bing background images.

.PARAMETER Resolution
Sets the native resolution of the retrieved images. Only some resolutions are
natively supported by Bing, and if the resolution is not supported, this script
will most likely fail. Get-BingBackgroundImages doesn't convert images, when
the -Resolution parameter is defined - it only tries to retrieve already available
images in the specified resolution from Bing. The format is [width]x[height] without
the opening or closing brackets. Valid -Resolution parameter values include, for
instance:

    1366x768
    1920x1080
    1920×1200
    1080x1920

.PARAMETER Subfolder
with aliases -SubfolderForThePortraitPictures, -SubfolderForTheVerticalPictures
and -SubfolderName. Specifies the name of the subfolder, where the new portrait
pictures are to be saved. The value for the -Subfolder parameter should be a plain
directory name (omitting the path and the parent directories). The default value is
"Vertical", which is used if no value for the -Subfolder parameter is defined in the
command launching Get-BingBackgroundImages. For best results in iterative usage of
Get-BingBackgroundImages, the default value should remain constant and be set
according to the prevailing conditions at line 23. Furthermore, if the directory
name includes space characters, please enclose the directory name in quotation marks
(single or double).

.PARAMETER Log
If the -Log parameter is added to the command launching Get-BingBackgroundImages,
a log file creation/updating procedure is initiated, if new images are downloaded.
The log file (bing_log.csv) is created or updated at the path defined with the
-Destination parameter. If the CSV log file seems to already exist, new data will
be appended to the end of that file. The log file contains over 30 image metadata
properties. For instance, the 'Custom URL' column lists a link to the homepage of
the picture at http://www.bing.com/gallery containing additional information about
the image and a detailed description.

A rather peculiar append procedure is used instead of the native -Append parameter
of the Export-Csv cmdlet for ensuring, that the CSV file will not contain any
additional quotation marks(, which might mess up the layout in some scenarios).

.PARAMETER IncludePortrait
with an alias -Include. If the -IncludePortrait parameter is used in the command
launching Get-BingBackgroundImages, and if the primary resolution set with the
-Resolution parameter seems to indicate that a horizontal (landscape) image is being
aquired, Get-BingBackgroundImages tries to download a particular horizontal
(landscape) image in the portrait format as well (with swapped width and height
values) along with the "normally" downloaded horizontal (landscape) image.
The landscape images are placed in the folder set with the -Destination parameter,
and the portrait pictures are placed to a subfolder inside the -Destination main
folder. The default name of the subfolder ("Vertical") may be changed with the
-Subfolder parameter.

.OUTPUTS
Displays a summary of the actions in console. Writes the Bing XML-file (bing.xml)
to $env:temp. Displays the downloaded files in a pop-up window (Out-GridView).
Saves the retrieved images in locations defined with the -Destination and -Subfolder
parameters. Additionally, if the -Log parameter is included in the command launching
Get-BingBackgroundImages, a log file creation updating procedure is initiated
at the path defined with the -Destination variable.


    Default values (the log file (bing_log.csv) creation/updating procedure only occurs,
    if the -Log parameter is used and new files are found):


    "$env:temp\bing.xml"                                        : XML-file
    "$env:USERPROFILE\Pictures\Wallpapers\Bing"                 : The folder for landscape wallpapers
    "$env:USERPROFILE\Pictures\Wallpapers\Bing\bing_log.csv"    : CSV-file
    "$env:USERPROFILE\Pictures\Wallpapers\Bing\Vertical"        : The folder for portrait pictures


.NOTES
Please note that all the parameters can be used in one get Bing background images
command and that each of the parameters can be "tab completed" before typing them
fully (by pressing the [tab] key).

To see the Bing Homepage Gallery, which showcases images featured on the U.S. Bing
homepage during the past 5 years, please visit http://www.bing.com/gallery/. For
a mobile alternative on Windows 10 Mobile, Windows Phone 8.1 or Windows Phone 8,
please check out the "Picture of the Day" -app at the Microsoft Store, which
downloads the pictures in 1080p directly to the Picture Hub
https://www.microsoft.com/en-US/store/p/picture-of-the-day/9nblggh09qtf).

Please note that the Bing XML-file (bing.xml) is saved to $env:temp directory,
and that the destination folders for images are end-user settable in each get Bing
background images command with the -Destination and -Subfolder parameters.
The $env:temp variable points to the current temp folder. The default value of the
$env:temp variable is C:\Users\<username>\AppData\Local\Temp (i.e. each user account
has their own separate temp folder at path %USERPROFILE%\AppData\Local\Temp). To see
the current temp path, for instance a command

    [System.IO.Path]::GetTempPath()

may be used at the PowerShell prompt window [PS>]. To change the temp folder for
instance to C:\Temp, please, for example, follow the instructions at
http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html

    Homepage:           https://github.com/auberginehill/get-bing-background-images
    Short URL:          http://tinyurl.com/jjnuvqr
    Version:            1.0

.EXAMPLE
./Get-BingBackgroundImages.ps1
Runs the script. Please notice to insert ./ or .\ before the script name. Tries to
download the Bing background image currently featured on Bing search engine homepage,
since no values for the -NumberOfImages or -Index parameters were defined. Saves
the downloaded Bing XML-file (bing.xml) to $env:temp. Queries the default
-Destination folder "$env:USERPROFILE\Pictures\Wallpapers\Bing" for existing files
and compares the filenames found in that folder to the filenames listed in the Bing
XML-file (bing.xml), and downloads all the files listed in the Bing XML-file
(bing.xml), which don't exist at the -Destination folder. Uses the default
market ("en-US") and the default resolution ("1920x1080") for the type of image to
download and saves the image to the default -Destination folder
("$env:USERPROFILE\Pictures\Wallpapers\Bing"). A pop-up window listing the new
files will open, if new files were downloaded.

.EXAMPLE
help ./Get-BingBackgroundImages -Full
Displays the help file.

.EXAMPLE
./Get-BingBackgroundImages.ps1 -Market en-UK -Index 0 -NumberOfImages 7 -Resolution 1920x1080 -Log -IncludePortrait
Tries to download seven latest background images from Bing (today's picture and six
previous pictures) in the landscape and portrait formats. If some "root" files 
listed in the Bing XML-file (bing.xml) seem not to exist at the default -Destination
folder in 1920x1080 resolution or at the default -Subfolder ("Vertical") in 
1080x1920 resolution, Get-BingBackgroundImages retrieves those images from the en-UK
Bing market. If new images were obtained, writes or updates a log file (bing_log.csv)
at the default -Destination folder, and a pop-up window listing the new files will open.

.EXAMPLE
./Get-BingBackgroundImages.ps1 -Index 7 -NumberOfImages 8 -Destination C:\Users\Dropbox\ -Subfolder Mobile -IncludePortrait
Tries to download eight pictures from Bing between the dates two weeks ago and one
week ago (including the start and end dates) in the landscape and portrait formats.
The actual download procedure is done "backwards" i.e. the newest photo is downloaded
first. If some files listed in the Bing XML-file (bing.xml) seem not to exist at
the -Destination folder ("C:\Users\Dropbox\") in the default (1920x1080) resolution
or at the -Subfolder ("Mobile") in 1080x1920 resolution, retrieves those images from
the default en-US Bing market, and a pop-up window listing the new files will open.
Since the path or the subfolder name doesn't contain any space characters, they
don't need to be enveloped with quotation marks.

.EXAMPLE
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
This command is altering the Windows PowerShell rights to enable script execution
in the default (LocalMachine) scope, and defines the conditions under which Windows
PowerShell loads configuration files and runs scripts in general. In Windows Vista
and later versions of Windows, for running commands that change the execution policy
of the LocalMachine scope, Windows PowerShell has to be run with elevated rights
(Run as Administrator). The default policy of the default (LocalMachine) scope is
"Restricted", and a command "Set-ExecutionPolicy Restricted" will "undo" the changes
made with the original example above (had the policy not been changed before...).
Execution policies for the local computer (LocalMachine) and for the current user
(CurrentUser) are stored in the registry (at for instance the
HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy key), and remain
effective until they are changed again. The execution policy for a particular session
(Process) is stored only in memory, and is discarded when the session is closed.


    Parameters:

    Restricted      Does not load configuration files or run scripts, but permits
                    individual commands. Restricted is the default execution policy.

    AllSigned       Scripts can run. Requires that all scripts and configuration
                    files be signed by a trusted publisher, including the scripts
                    that have been written on the local computer. Risks running
                    signed, but malicious, scripts.

    RemoteSigned    Requires a digital signature from a trusted publisher on scripts
                    and configuration files that are downloaded from the Internet
                    (including e-mail and instant messaging programs). Does not
                    require digital signatures on scripts that have been written on
                    the local computer. Permits running unsigned scripts that are
                    downloaded from the Internet, if the scripts are unblocked by
                    using the Unblock-File cmdlet. Risks running unsigned scripts
                    from sources other than the Internet and signed, but malicious,
                    scripts.

    Unrestricted    Loads all configuration files and runs all scripts.
                    Warns the user before running scripts and configuration files
                    that are downloaded from the Internet. Not only risks, but
                    actually permits, eventually, running any unsigned scripts from
                    any source. Risks running malicious scripts.

    Bypass          Nothing is blocked and there are no warnings or prompts.
                    Not only risks, but actually permits running any unsigned scripts
                    from any source. Risks running malicious scripts.

    Undefined       Removes the currently assigned execution policy from the current
                    scope. If the execution policy in all scopes is set to Undefined,
                    the effective execution policy is Restricted, which is the
                    default execution policy. This parameter will not alter or
                    remove the ("master") execution policy that is set with a Group
                    Policy setting.
    __________
    Notes: 	      - Please note that the Group Policy setting "Turn on Script Execution"
                    overrides the execution policies set in Windows PowerShell in all
                    scopes. To find this ("master") setting, please, for example, open
                    the Local Group Policy Editor (gpedit.msc) and navigate to
                    Computer Configuration > Administrative Templates >
                    Windows Components > Windows PowerShell.

                  - The Local Group Policy Editor (gpedit.msc) is not available in any
                    Home or Starter edition of Windows.

                  - Group Policy setting "Turn on Script Execution":

               	    Not configured                                          : No effect, the default
                                                                               value of this setting
                    Disabled                                                : Restricted
                    Enabled - Allow only signed scripts                     : AllSigned
                    Enabled - Allow local scripts and remote signed scripts : RemoteSigned
                    Enabled - Allow all scripts                             : Unrestricted


For more information, please type "Get-ExecutionPolicy -List", "help Set-ExecutionPolicy -Full",
"help about_Execution_Policies" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx
or http://go.microsoft.com/fwlink/?LinkID=135170.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-BingBackgroundImages.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent
-NoClobber mode built into it, so that the procedure will halt, if overwriting (replacing
the contents) of an existing file is about to happen. Overwriting a file with the New-Item
cmdlet requires using the Force. If the path name and/or the filename includes space
characters, please enclose the whole -Path parameter value in quotation marks (single or
double):

    New-Item -ItemType File -Path "C:\Folder Name\Get-BingBackgroundImages.ps1"

For more information, please type "help New-Item -Full".

.LINK
https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27
http://powershell.com/cs/media/p/7476.aspx
http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx
http://powershell.com/cs/media/p/32274.aspx ( http://web.archive.org/web/20150110213118/http://powershell.com/cs/media/p/32274.aspx )
https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/
https://msdn.microsoft.com/en-us/library/system.uri(v=vs.110).aspx
http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string
http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
https://blog.hyperexpert.com/powershell-script-to-download-bings-daily-image-as-a-wallpaper/
http://stackoverflow.com/questions/10639914/is-there-a-way-to-get-bings-photo-of-the-day
https://www.codeproject.com/articles/151937/bing-image-download
https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27

#>
