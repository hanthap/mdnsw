$d = Import-csv  "$unzippedRoot\Case_Note__c.csv" | where Id -eq 'a0Q3b00000KhBUSEA3' # oddly UTF8 seems to be the default, which is good except for apostrophes

$d.Notes__c | Out-File -FilePath "$unzippedRoot\example.html"

# Iteratively remove <img>...</img> (from bottom to top) until the resulting HTML string is less than X bytes in length
function Remove-Images {

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      $PropertyList=@('Notes__c'),
      $MaxLength = 32768,
      [string] $ReplaceWithText = '<H3 style="color: red">[Embedded image has been removed due to Maica 32k character limit]</H3>',
      [System.Text.RegularExpressions.RegexOptions] $RegexOptions = ([System.Text.RegularExpressions.RegexOptions]::RightToLeft)
      )

begin {


    $rxFooter = New-Object -TypeName regex -ArgumentList '\<img alt="email footer".*\<\/img\>', ($RegexOptions)
    $rxAny = New-Object -TypeName regex -ArgumentList '\<img.*\<\/img\>', ($RegexOptions)
    [regex] $rxHidden = '�'

}

process {


    foreach( $PropName in $PropertyList ) {
            $s = $obj.$PropName
            #remove all "email footer" images with no warning text
            $s.Replace( '<span.*></span>', '' ) # junk span tags
            $s = $rxFooter.Replace( $s, '' ) 
            $s = $rxHidden.Replace( $s, '' ) 

            #  insert warning message in place of the last (up to N remaining) images
            for ( $i=0; $i -lt 10 -and $s.Length -gt $MaxLength; $i++ ) {
                #remove all footer images 
                 $s = $rxAny.Replace( $s, $ReplaceWithText, 1 ) 
            }
            $obj.$PropName = $s
        }


    $obj

    }

}

$e = $d | Remove-Images -Verbose



$e.Notes__c | Out-File -FilePath "$unzippedRoot\example.html" -Encoding utf8


<span style="font-family: Calibri,sans-serif;"> </span>

[int][char]"�"


$s = ' <font face="Calibri, sans-serif"><span style="font-size: 14.6667px;">PC to Deepesh to see if he got capacity to take new participants to which he yes he does.<br>Informed Deepesh briefly re Ray and what we are looking for. Informed Deepesh that I will email Ray&#39;s plan details along with plan manager details shortly.</span></font><br><br><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"><span style="color: black;">From: Ganesh Kakani &lt;<a href="mailto:ganesh.kakani@mdnsw.org.au" style="color: blue; text-decoration: underline;" target="_blank">ganesh.kakani@mdnsw.org.au</a>&gt; </span></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"><span style="color: black;">Date: 6/4/21 11:36 am (GMT+10:00) </span></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"><span style="color: black;">To: DS &lt;<a href="mailto:deepesh.shresthaa@gmail.com" style="color: blue; text-decoration: underline;" target="_blank">deepesh.shresthaa@gmail.com</a>&gt; </span></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"><span style="color: black;">Subject: Ray Grasso </span></span></span><br><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Good Morning Deepesh,</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Thank you for talking time to talk to me re Ray Grasso on Thursday 01 April 2021.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">As informed Ray is looking for Physio Therapist to assist him with maintain his current mobility and strength.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">To start with can you please able to meet Ray to go through what he is looking for and what you can able to assist Ray with.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Can you please able to do a service agreement for 3 hours for initial assessment and recommendations.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Please find Ray?s plan details below for your reference.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"><H3 style="color: red">[Embedded image has been removed due to Maica 32k character limit]</H3></span></span><br><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Ray?s plan is plan managed by Zest Care so can you please send the invoices to <a href="mailto:planmanager@zestcare.net.au" style="color: blue; text-decoration: underline;" target="_blank">planmanager@zestcare.net.au</a></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">As discussed and informed Ray prefers to be contacted after 1 pm as he struggles sleeping at night so he will be resting during the day time.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Can you please contact him on 0450 084 480. If you can?t reach him? can you text him.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Please let me know if you needs any further information.</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;"></span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Best Regards,</span></span><br><span style="font-size: 11pt;"><span style="font-family: Calibri,sans-serif;">Ganesh Kakani</span></span><br><br>'

$s.length

$t = $s.Replace( '<span.*></span>', '' )


$t.length

 $t | Out-File -FilePath "$unzippedRoot\example.html" -Encoding utf8