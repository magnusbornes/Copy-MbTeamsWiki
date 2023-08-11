# Copy-MbTeamsWiki

This PowerShell script lets you copy any Teams Wiki from Standard og Private Channels to any other Channel.  
(It might work with Shared Channels to, but I never tested that.)

## How to use:
______________
    1.  Open Teams.
    2.  Right-click, and copy link to the tab in the Channel where the Wiki is located.
            This is your sourceUrl.
    3.  Either, Create a new Teams Wiki tab in the target Channel and copy the link, or copy the link to the Channel (Right-click in left menu!).
            This is your targetUrl. (I recommend doing the former option...)
        
    $SourceUrl = "https://teams.microsoft.com/l/channel/19%3Addb24db9b083441871a99e09148f6dc4%40thread.tacv2/tab%3A%3A95527d65-6cef-4b31-9f68-c21e396444df?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
    
    $TargetUrl = "https://teams.microsoft.com/l/channel/19%3a12fac4bb21db4d6e8e616d0123489a93%40thread.tacv2/GR01-renamed?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
    
    Copy-MbTeamWiki.ps1 -SourceUrl "$"SourceUrl -TargetUrl "$"TargetUrl
    
    This will copy Wiki from the provided source tab, to the provided target tab.
    
_______________
As Teams Wiki is being depricated, this script has purose for a limited amout of time.  
Therefore I have not spent a lot of time optimizing the script with parallelization and what not.

The script is pretty messy, but hey; if it works – don't fix it, am I right?
I also only have a couple of years of experience with PowerShell, so bear with me.  

## "Known" issues:  
-   If image filename already exists, the copy will fail, and you will have to run the script all over, or skip that image.
-   Some of the regex patterns are pretty jank, and they may fail if you have something unusual in your Wiki, that I haven't tested.
-   Error handling could absolutely be better, but it's not like any enduser is going to run this script anyway.
-   Host output could also be better.


Created by me, Magnus Børnes @ Norwegian University of Science and Technology, August 2023.