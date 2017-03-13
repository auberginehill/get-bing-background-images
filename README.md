<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V
     Visual Studio Code: To crop the tailing end space characters out, please use the key combination Ctrl + A Ctrl + K Ctrl + X (Formerly Ctrl + Shift + X)
     Visual Studio Code: To improve the formatting of HTML code, press Shift + Alt + F and the selected area will be reformatted in a html file.
     Visual Studio Code shortcuts: http://code.visualstudio.com/docs/customization/keybindings (or https://aka.ms/vscodekeybindings)
     Visual Studio Code shortcut PDF (Windows): https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf


   _____      _          ____  _             ____             _                                   _ _____
  / ____|    | |        |  _ \(_)           |  _ \           | |                                 | |_   _|
 | |  __  ___| |_ ______| |_) |_ _ __   __ _| |_) | __ _  ___| | ____ _ _ __ ___  _   _ _ __   __| | | |  _ __ ___   __ _  __ _  ___  ___
 | | |_ |/ _ \ __|______|  _ <| | '_ \ / _` |  _ < / _` |/ __| |/ / _` | '__/ _ \| | | | '_ \ / _` | | | | '_ ` _ \ / _` |/ _` |/ _ \/ __|
 | |__| |  __/ |_       | |_) | | | | | (_| | |_) | (_| | (__|   < (_| | | | (_) | |_| | | | | (_| |_| |_| | | | | | (_| | (_| |  __/\__ \
  \_____|\___|\__|      |____/|_|_| |_|\__, |____/ \__,_|\___|_|\_\__, |_|  \___/ \__,_|_| |_|\__,_|_____|_| |_| |_|\__,_|\__, |\___||___/
                                        __/ |                      __/ |                                                   __/ |
                                       |___/                      |___/                                                   |___/                               -->


## Get-BingBackgroundImages.ps1

<table>
    <tr>
        <td style="padding:6px"><strong>OS:</strong></td>
        <td colspan="2" style="padding:6px">Windows</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Type:</strong></td>
        <td colspan="2" style="padding:6px">A Windows PowerShell script</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Language:</strong></td>
        <td colspan="2" style="padding:6px">Windows PowerShell</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Description:</strong></td>
        <td colspan="2" style="padding:6px">
            <p>
                Get-BingBackgroundImages downloads a Bing <a href="http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=en-US">XML-file</a> (<code>bing.xml</code>) from Bing <code>http://www.bing.com/HPImageArchive.aspx?format=xml&amp;idx=[Index]&amp;n=[NumberOfImages]&amp;mkt=[Market]</code> containing info about the queried daily background images, which were recently featured on the Bing search engine <a href="http://www.bing.com/">homepage</a> as a daily changing background image. Microsoft's Bing search engine is famous for using hand-picked beautiful imagery on their main landing page. The daily background images regularly depict stunning landmarks, inspiring nature shots and high quality pictures of people and cultures from around the world.</p>
            <p>
                After determining the save locations, which are set with <code>-Destination</code> and <code>-Subfolder</code> parameters, Get-BingBackgroundImages compares the filenames found in the destination folders to the filenames listed in the XML-file (<code>bing.xml</code>) and tries to determine, whether downloading an image is actually needed or not. Get-BingBackgroundImages tries to avoid any unnecessary downloading and overwriting of the existing files.</p>
            <p>
                By using the Bing API with AJAX calls it seems to be possible to retrieve up to 14 days old pictures and up to 8 images at a time. To retrieve all currently available Bing background images in the default 1920x1080 resolution along with the 1080x1920 portrait variants as well from the default Bing market (<code>en-US</code>) with Get-BingBackgroundImages and to create/update a log file (<code>bing_log.csv</code>), please use two separate commands launching Get-BingBackgroundImages:</p>
            <p>
                <ul>
                    <code>.\Get-BingBackgroundImages.ps1 -Index 0 -NumberOfImages 7 -Log -IncludePortrait</code><br />
                    <code>.\Get-BingBackgroundImages.ps1 -Index 7 -NumberOfImages 8 -Log -IncludePortrait</code>
                </ul></p>
            <p>
                Please note that Get-BingBackgroundImages proceeds towards the history, when handling multiple images. Please also note that if any of the individual parameter values include space characters, the individual value should be enclosed in quotation marks (single or double) so that PowerShell can interpret the command correctly. This script is based on Jeffrey's script "<a href="https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27">Get-BingImage.ps1</a>".</p>
        </td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Homepage:</strong></td>
        <td colspan="2" style="padding:6px"><a href="https://github.com/auberginehill/get-bing-background-images">https://github.com/auberginehill/get-bing-background-images</a>
            <br />Short URL: <a href="http://tinyurl.com/jjnuvqr">http://tinyurl.com/jjnuvqr</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Version:</strong></td>
        <td colspan="2" style="padding:6px">1.0</td>
    </tr>
    <tr>
        <td rowspan="5" style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">Emojis:</td>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Jeffrey:</td>
        <td style="padding:6px"><a href="https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27">Get-BingImage.ps1</a></td>
    </tr>
    <tr>
        <td style="padding:6px">clayman2:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">ps1:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx">Test Internet connection</a> (or one of the <a href="https://web.archive.org/web/20110612212629/http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Tobias Weltner:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/media/p/32274.aspx">PowerTips Monthly vol 10 March 2014</a>, Read raw web page content (or one of the <a href="http://web.archive.org/web/20150110213118/http://powershell.com/cs/media/p/32274.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Downloads:</strong></td>
        <td colspan="2" style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/get-bing-background-images/master/Get-BingBackgroundImages.ps1">Get-BingBackgroundImages.ps1</a>. Or <a href="https://github.com/auberginehill/get-bing-background-images/archive/master.zip">everything as a .zip-file</a>.</td>
    </tr>
</table>




### Screenshot

<ol><ol><ol>
<img class="screenshot" title="screenshot" alt="screenshot" height="80%" width="80%" src="https://raw.githubusercontent.com/auberginehill/get-bing-background-images/master/Get-BingBackgroundImages.png">
</ol></ol></ol>




### Parameters

<table>
    <tr>
        <th>:triangular_ruler:</th>
        <td style="padding:6px">
            <ul>
                <li>
                    <h5>Parameter <code>-Destination</code></h5>
                    <p>with aliases <code>-Path</code>, <code>-Output</code> and <code>-OutputFolder</code>. Specifies the primary folder, where the acquired new Bing background images are to be saved, and defines the default location to be used with the <code>-Log</code> parameter (<code>bing_log.csv</code>), and also sets the parent directory for the <code>-Subfolder</code> parameter. The default save location for the horizontal (landscape) Bing background images is <code>"$env:USERPROFILE\Pictures\Wallpapers\Bing"</code>, which will be used, if no value for the <code>-Destination</code> parameter is defined in the command launching Get-BingBackgroundImages. For best results in iterative usage of this script, the default value should remain constant and be set according to the prevailing conditions (at line 13). The value for the <code>-Destination</code> parameter should be a valid file system path pointing to a directory (a full path of a folder such as <code>C:\Users\Dropbox\</code>). Furthermore, if the path includes space characters, please enclose the path in quotation marks (single or double).</p>
                </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>
                        <h5>Parameter <code>-Market</code></h5>
                        <p>Sets the Bing market, from where the images are downloaded. The valid values are: <code>en-US</code>, <code>zh-CN</code>, <code>ja-JP</code>, <code>en-AU</code>, <code>en-UK</code>, <code>de-DE</code>, <code>en-NZ</code> and <code>en-CA</code>. If no value for the <code>-Market</code> parameter is defined in the command launching Get-BingBackgroundImages, the default value "<code>en-US</code>" will be used.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Index</code></h5>
                        <p>Sets the start date, from which point towards the history should the images be retrieved. 0 denominates the current day, 1 denominates yesterday, 2 denominates the day before yesterday etc. It should be noted that Get-BingBackgroundImages always counts "<dfn>backwads</dfn>" (towards the history), when retrieving multiple Bing background images, and that by default, Get-BingBackgroundImages starts to count the images from the current day.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-NumberOfImages</code></h5>
                        <p>with aliases <code>-Number</code>, <code>-Amount</code> and <code>-Images</code>. Sets the number of images to return. n = 1 would return only one, n = 2 would return two, and so on up to eight. By default Get-BingBackgroundImages retrieves one image. It should also be noted that Get-BingBackgroundImages always counts "<dfn>backwads</dfn>" (towards the history), when retrieving Bing background images.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Resolution</code></h5>
                        <p>Sets the native resolution of the retrieved images. Only some resolutions are natively supported by Bing, and if the resolution is not supported, this script will most likely fail. Get-BingBackgroundImages doesn't convert images, when the <code>-Resolution</code> parameter is defined – it only tries to retrieve already available images in the specified resolution from Bing. The format is [width]x[height] without the opening or closing brackets. Valid <code>-Resolution</code> parameter values include, for instance:</p>
                        <p>
                            <ul>
                                <code>1366x768</code>
                                <br /><code>1920x1080</code>
                                <br /><code>1920×1200</code>
                                <br /><code>1080x1920</code>
                            </ul>
                        </p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Subfolder</code></h5>
                        <p>with aliases <code>-SubfolderForThePortraitPictures</code>, <code>-SubfolderForTheVerticalPictures</code> and <code>-SubfolderName</code>. Specifies the name of the subfolder, where the new portrait pictures are to be saved. The value for the <code>-Subfolder</code> parameter should be a plain directory name (omitting the path and the parent directories). The default value is "<code>Vertical</code>", which is used if no value for the <code>-Subfolder</code> parameter is defined in the command launching Get-BingBackgroundImages. For best results in iterative usage of Get-BingBackgroundImages, the default value should remain constant and be set according to the prevailing conditions at line 23. Furthermore, if the directory name includes space characters, please enclose the directory name in quotation marks (single or double).</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Log</code></h5>
                        <p>If the <code>-Log</code> parameter is added to the command launching Get-BingBackgroundImages, a log file creation/updating procedure is initiated, if new images are downloaded. The log file (<code>bing_log.csv</code>) is created or updated at the path defined with the <code>-Destination</code> parameter. If the CSV log file seems to already exist, new data will be appended to the end of that file. The log file contains over 30 image metadata properties. For instance, the '<code>Custom URL</code>' column lists a link to the homepage of the picture at http://www.bing.com/gallery containing additional information about the image and a detailed description.</p>
                        <p>A rather peculiar append procedure is used instead of the native -Append parameter of the <code>Export-Csv</code> cmdlet for ensuring, that the CSV file will not contain any additional quotation marks(, which might mess up the layout in some scenarios).</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-IncludePortrait</code></h5>
                        <p>with an alias <code>-Include</code>. If the <code>-IncludePortrait</code> parameter is used in the command launching Get-BingBackgroundImages, and if the primary resolution set with the <code>-Resolution</code> parameter seems to indicate that a horizontal (landscape) image is being aquired, Get-BingBackgroundImages tries to download a particular horizontal (landscape) image in the portrait format as well (with swapped width and height values) along with the "<dfn>normally</dfn>" downloaded horizontal (landscape) image. The landscape images are placed in the folder set with the <code>-Destination</code> parameter, and the portrait pictures are placed to a subfolder inside the <code>-Destination</code> main folder. The default name of the subfolder ("<code>Vertical</code>") may be changed with the <code>-Subfolder</code> parameter.</p>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>Displays a summary of the actions in console. Writes the Bing XML-file (<code>bing.xml</code>) to <code>$env:temp</code>. Displays the downloaded files in a pop-up window (<code>Out-GridView</code>). Saves the retrieved images in locations defined with the <code>-Destination</code> and <code>-Subfolder</code> parameters. Additionally, if the <code>-Log</code> parameter is included in the command launching Get-BingBackgroundImages, a log file creation updating procedure is initiated  at the path defined with the <code>-Destination</code> variable.</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <ol>
                    <p>Default values (the log file (<code>bing_log.csv</code>) creation/updating procedure only occurs, if the <code>-Log</code> parameter is used and new files are found):</p>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Path</strong></td>
                                <td style="padding:6px"><strong>Type</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$env:temp\bing.xml"</code></td>
                                <td style="padding:6px">XML-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$env:USERPROFILE\Pictures\Wallpapers\Bing"</code></td>
                                <td style="padding:6px">The folder for landscape wallpapers</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$env:USERPROFILE\Pictures\Wallpapers\Bing\bing_log.csv"</code></td>
                                <td style="padding:6px">CSV-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>"$env:USERPROFILE\Pictures\Wallpapers\Bing\Vertical"</code></td>
                                <td style="padding:6px">The folder for portrait pictures</td>
                            </tr>
                        </table>
                    </p>
                </ol>
            </ul>
        </td>
    </tr>
</table>




### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that all the parameters can be used in one get Bing background images command and that each of the parameters can be "tab completed" before typing them fully (by pressing the <code>[tab]</code> key).</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>To see the Bing Homepage Gallery, which showcases images featured on the U.S. Bing homepage during the past 5 years, please visit <a href="http://www.bing.com/gallery/">http://www.bing.com/gallery/</a>. For a mobile alternative on Windows 10 Mobile, Windows Phone 8.1 or Windows Phone 8, please check out the <a href="https://www.microsoft.com/en-US/store/p/picture-of-the-day/9nblggh09qtf">Picture of the Day</a> -app at the Microsoft Store, which downloads the pictures in 1080p directly to the Picture Hub.</li>
                </p>
                <p>
                    <li>Please note that the Bing XML-file (<code>bing.xml</code>) is saved to <code>$env:temp</code> directory, and that the destination folders for images are end-user settable in each get Bing background images command with the <code>-Destination</code> and <code>-Subfolder</code> parameters. The <code>$env:temp</code> variable points to the current temp folder. The default value of the <code>$env:temp</code> variable is <code>C:\Users\&lt;username&gt;\AppData\Local\Temp</code> (i.e. each user account has their own separate temp folder at path <code>%USERPROFILE%\AppData\Local\Temp</code>). To see the current temp path, for instance a command
                    <br />
                    <br /><code>[System.IO.Path]::GetTempPath()</code>
                    <br />
                    <br />may be used at the PowerShell prompt window <code>[PS&gt;]</code>. To change the temp folder for instance to <code>C:\Temp</code>, please, for example, follow the instructions at <a href="http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html">Temporary Files Folder - Change Location in Windows</a>, which in essence are something along the lines:
                        <ol>
                           <li>Right click Computer icon and select Properties (or select Start → Control Panel → System. On Windows 10 this instance may also be found by right clicking Start and selecting Control Panel → System... or by pressing <code>[Win-key]</code> + X and selecting Control Panel → System). On the window with basic information about the computer...</li>
                           <li>Click on Advanced system settings on the left panel and select Advanced tab on the "System Properties" pop-up window.</li>
                           <li>Click on the button near the bottom labeled Environment Variables.</li>
                           <li>In the topmost section, which lists the User variables, both TMP and TEMP may be seen. Each different login account is assigned its own temporary locations. These values can be changed by double clicking a value or by highlighting a value and selecting Edit. The specified path will be used by Windows and many other programs for temporary files. It's advisable to set the same value (a directory path) for both TMP and TEMP.</li>
                           <li>Any running programs need to be restarted for the new values to take effect. In fact, probably Windows itself needs to be restarted for it to begin using the new values for its own temporary files.</li>
                        </ol>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Examples

<table>
    <tr>
        <th>:book:</th>
        <td style="padding:6px">To open this code in Windows PowerShell, for instance:</td>
   </tr>
   <tr>
        <th></th>
        <td style="padding:6px">
            <ol>
                <p>
                    <li><code>./Get-BingBackgroundImages.ps1</code><br />
                    Runs the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name. Tries to download the Bing background image currently featured on Bing search engine homepage, since no values for the <code>-NumberOfImages</code> or <code>-Index</code> parameters were defined. Saves the downloaded Bing XML-file (<code>bing.xml</code>) to <code>$env:temp</code>. Queries the default <code>-Destination</code> folder <code>"$env:USERPROFILE\Pictures\Wallpapers\Bing"</code> for existing files and compares the filenames found in that folder to the filenames listed in the Bing XML-file (<code>bing.xml</code>), and downloads all the files listed in the Bing XML-file (<code>bing.xml</code>), which don't exist at the <code>-Destination</code> folder. Uses the default market ("<code>en-US</code>") and the default resolution ("<code>1920x1080</code>") for the type of image to download and saves the image to the default <code>-Destination</code> folder (<code>"$env:USERPROFILE\Pictures\Wallpapers\Bing"</code>). A pop-up window listing the new files will open, if new files were downloaded.</li>
                </p>
                <p>
                    <li><code>help ./Get-BingBackgroundImages -Full</code><br />
                    Displays the help file.</li>
                </p>
                <p>
                    <li><code>./Get-BingBackgroundImages.ps1 -Market en-UK -Index 0 -NumberOfImages 7 -Resolution 1920x1080 -Log -IncludePortrait</code><br />
                    Tries to download seven latest background images from Bing (today's picture and six previous pictures) in the landscape and portrait formats. If some "<dfn>root</dfn>" files listed in the Bing XML-file (<code>bing.xml</code>) seem not to exist at the default <code>-Destination</code> folder in 1920x1080 resolution or at the default <code>-Subfolder</code> ("<code>Vertical</code>") in 1080x1920 resolution, retrieves those images from the en-UK Bing market. If new images were obtained, writes or updates a log file (<code>bing_log.csv</code>) at the default <code>-Destination</code> folder, and a pop-up window listing the new files will open.</li>
                </p>
                <p>
                    <li><code>./Get-BingBackgroundImages.ps1 -Index 7 -NumberOfImages 8 -Destination C:\Users\Dropbox\ -Subfolder Mobile -IncludePortrait</code><br />
                    Tries to download eight pictures from Bing between the dates two weeks ago and one week ago (including the start and end dates) in the landscape and portrait formats. The actual download procedure is done "backwards" i.e. the newest photo is downloaded first. If some files listed in the Bing XML-file (<code>bing.xml</code>) seem not to exist at the <code>-Destination</code> folder ("<code>C:\Users\Dropbox\</code>") in the default (1920x1080) resolution or at the <code>-Subfolder</code> ("Mobile") in 1080x1920 resolution, retrieves those images from the default <code>en-US</code> Bing market, and a pop-up window listing the new files will open. Since the path or the subfolder name doesn't contain any space characters, they don't need to be enveloped with quotation marks.</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution in the default (<code>LocalMachine</code>) scope, and defines the conditions under which Windows PowerShell loads configuration files and runs scripts in general. In Windows Vista and later versions of Windows, for running commands that change the execution policy of the <code>LocalMachine</code> scope, Windows PowerShell has to be run with elevated rights (<dfn>Run as Administrator</dfn>). The default policy of the default (<code>LocalMachine</code>) scope is "<code>Restricted</code>", and a command "<code>Set-ExecutionPolicy Restricted</code>" will "<dfn>undo</dfn>" the changes made with the original example above (had the policy not been changed before...). Execution policies for the local computer (<code>LocalMachine</code>) and for the current user (<code>CurrentUser</code>) are stored in the registry (at for instance the <code>HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy</code> key), and remain effective until they are changed again. The execution policy for a particular session (<code>Process</code>) is stored only in memory, and is discarded when the session is closed.</p>
                        <p>Parameters:
                            <ul>
                                <table>
                                    <tr>
                                        <td style="padding:6px"><code>Restricted</code></td>
                                        <td colspan="2" style="padding:6px">Does not load configuration files or run scripts, but permits individual commands. <code>Restricted</code> is the default execution policy.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>AllSigned</code></td>
                                        <td colspan="2" style="padding:6px">Scripts can run. Requires that all scripts and configuration files be signed by a trusted publisher, including the scripts that have been written on the local computer. Risks running signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>RemoteSigned</code></td>
                                        <td colspan="2" style="padding:6px">Requires a digital signature from a trusted publisher on scripts and configuration files that are downloaded from the Internet (including e-mail and instant messaging programs). Does not require digital signatures on scripts that have been written on the local computer. Permits running unsigned scripts that are downloaded from the Internet, if the scripts are unblocked by using the <code>Unblock-File</code> cmdlet. Risks running unsigned scripts from sources other than the Internet and signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Unrestricted</code></td>
                                        <td colspan="2" style="padding:6px">Loads all configuration files and runs all scripts. Warns the user before running scripts and configuration files that are downloaded from the Internet. Not only risks, but actually permits, eventually, running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Bypass</code></td>
                                        <td colspan="2" style="padding:6px">Nothing is blocked and there are no warnings or prompts. Not only risks, but actually permits running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Undefined</code></td>
                                        <td colspan="2" style="padding:6px">Removes the currently assigned execution policy from the current scope. If the execution policy in all scopes is set to <code>Undefined</code>, the effective execution policy is <code>Restricted</code>, which is the default execution policy. This parameter will not alter or remove the ("<dfn>master</dfn>") execution policy that is set with a Group Policy setting.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px; border-top-width:1px; border-top-style:solid;"><span style="font-size: 95%">Notes:</span></td>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li><span style="font-size: 95%">Please note that the Group Policy setting "<code>Turn on Script Execution</code>" overrides the execution policies set in Windows PowerShell in all scopes. To find this ("<dfn>master</dfn>") setting, please, for example, open the Local Group Policy Editor (<code>gpedit.msc</code>) and navigate to Computer Configuration → Administrative Templates → Windows Components → Windows PowerShell.</span></li>
                                            </ul>
                                        </td>
                                    </tr>
                                    <tr>
                                        <th></th>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li><span style="font-size: 95%">The Local Group Policy Editor (<code>gpedit.msc</code>) is not available in any Home or Starter edition of Windows.</span></li>
                                                <ol>
                                                    <p>
                                                        <table>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><strong>Group Policy Setting</strong> "<code>Turn&nbsp;on&nbsp;Script&nbsp;Execution</code>"</td>
                                                                <td style="padding:6px; font-size: 85%"><strong>PowerShell Equivalent</strong> (concerning all scopes)</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Not configured</code></td>
                                                                <td style="padding:6px; font-size: 85%">No effect, the default value of this setting</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Disabled</code></td>
                                                                <td style="padding:6px; font-size: 85%"><code>Restricted</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow only signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>AllSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow local scripts and remote signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>RemoteSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> – Allow all scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>Unrestricted</code></td>
                                                            </tr>
                                                        </table>
                                                    </p>
                                                </ol>
                                            </ul>
                                        </td>
                                    </tr>
                                </table>
                            </ul>
                        </p>
                    <p>For more information, please type "<code>Get-ExecutionPolicy -List</code>", "<code>help Set-ExecutionPolicy -Full</code>", "<code>help about_Execution_Policies</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a> or <a href="http://go.microsoft.com/fwlink/?LinkID=135170">about_Execution_Policies</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Get-BingBackgroundImages.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>. If the path name and/or the filename includes space characters, please enclose the whole <code>-Path</code> parameter value in quotation marks (single or double):
                        <ol>
                            <br /><code>New-Item -ItemType File -Path "C:\Folder Name\Get-BingBackgroundImages.ps1"</code>
                        </ol>
                    <br />For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




### Contributing

<table>
    <tr>
        <th><img class="emoji" title="contributing" alt="contributing" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f33f.png"></th>
        <td style="padding:6px"><strong>Bugs:</strong></td>
        <td style="padding:6px">Bugs can be reported by creating a new <a href="https://github.com/auberginehill/get-bing-background-images/issues">issue</a>.</td>
    </tr>
    <tr>
        <th rowspan="2"></th>
        <td style="padding:6px"><strong>Feature Requests:</strong></td>
        <td style="padding:6px">Feature request can be submitted by creating a new <a href="https://github.com/auberginehill/get-bing-background-images/issues">issue</a>.</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Editing Source Files:</strong></td>
        <td style="padding:6px">New features, fixes and other potential changes can be discussed in further detail by opening a <a href="https://github.com/auberginehill/get-bing-background-images/pulls">pull request</a>.</td>
    </tr>
</table>




### www

<table>
    <tr>
        <th>:globe_with_meridians:</th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-bing-background-images">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="12"></th>
        <td style="padding:6px">Jeffrey: <a href="https://gist.github.com/jeffpatton1971/437b8487ae7e69ba4d27">Get-BingImage.ps1</a></td>
    </tr>
    <tr>
        <td style="padding:6px">clayman2: <a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">ps1: <a href="http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx">Test Internet connection</a> (or one of the <a href="https://web.archive.org/web/20110612212629/http://powershell.com/cs/blogs/tips/archive/2011/05/04/test-internet-connection.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Tobias Weltner: <a href="http://powershell.com/cs/media/p/32274.aspx">PowerTips Monthly vol 10 March 2014</a>, Read raw web page content (or one of the <a href="http://web.archive.org/web/20150110213118/http://powershell.com/cs/media/p/32274.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/">Remove Unwanted Quotation Marks from CSV Files by Using PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.uri(v=vs.110).aspx">Uri Class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string">Powershellv2 - remove last x characters from a string</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2">Append files using export-csv for PowerShell 2</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blog.hyperexpert.com/powershell-script-to-download-bings-daily-image-as-a-wallpaper/">PowerShell script to download Bing's daily image as a wallpaper</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/10639914/is-there-a-way-to-get-bings-photo-of-the-day">Is there a way to get Bing's photo of the day?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.codeproject.com/articles/151937/bing-image-download">Bing Image Download</a></td>
    </tr>
    <tr>
        <td style="padding:6px">ASCII Art: <a href="http://www.figlet.org/">http://www.figlet.org/</a> and <a href="http://www.network-science.de/ascii/">ASCII Art Text Generator</a></td>
    </tr>
</table>




### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/aa812bfa79fa19fbd880b97bdc22e2c1">Disable-Defrag</a></td>
    </tr>
    <tr>
        <th rowspan="27"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/firefox-customization-files">Firefox Customization Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ascii-table">Get-AsciiTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-battery-info">Get-BatteryInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Get-ComputerInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-culture-tables">Get-CultureTables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-directory-size">Get-DirectorySize</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-hash-value">Get-HashValue</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-programs">Get-InstalledPrograms</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-windows-updates">Get-InstalledWindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-powershell-aliases-table">Get-PowerShellAliasesTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/9c2f26146a0c9d3d1f30ef0395b6e6f5">Get-PowerShellSpecialFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-time-zone-table">Get-TimeZoneTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers">Get-Windows10LockScreenWallpapers</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/java-update">Java-Update</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-duplicate-files">Remove-DuplicateFiles</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-empty-folders">Remove-EmptyFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/13bb9f56dc0882bf5e85a8f88ccd4610">Remove-EmptyFoldersLite</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/176774de38ebb3234b633c5fbc6f9e41">Rename-Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/rock-paper-scissors">Rock-Paper-Scissors</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/toss-a-coin">Toss-a-Coin</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">Unzip-Silently</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-adobe-flash-player">Update-AdobeFlashPlayer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-mozilla-firefox">Update-MozillaFirefox</a></td>
    </tr>
</table>
