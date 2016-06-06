Upgrade Information - Proc Script Amendment

The current bundle has retired older proc scripts - meaning certain data types and files are no longer updated (win-procinfo, win-proclist, win-procperf).

If you are installing onto a fresh environment, you need do nothing.

If you already are running a previous version of the WindowsApp
AND
You have built custom workspaces/searches using the WindowsApp data 
THEN
You will need to check if your alerts/searches/workspaces will be affected by the transition. 

OPTION 1 - I WANT TO CHANGE MY ENVIRONMENT NOW

You can do this by:
1. Exporting an XML backup of your environment Workspaces and Searches
2. Search for the following:
  -  win-procinfo
  -  win-proclist
  -  win-procperf

You will probably find them in use in the following way:
  * | _type.equals(win-proclist)

  Simply amend the search to include the new data type like the following example the following

  * | _type.equals(win-proclist,win-process)

Then search your XML for the following

  -  Win_Processes.out
  -  Win_ProcInfo.out
  -  Win_ProcPerf.out

You will probably find them in use in the following way:
  * | _filename.include(Win_Processes.out)

  Simply amend the search to include the new data type like the following example the following

  * | _filename.include(Win_Processes.out,Win_ActiveProcess.out,Win_FullProcess.out)

As you will see, the data types have the same column names as the old data - therefore any searches relying on the data type will be backwards compatible so long as you include the new data in your search parameters. 

3. Once you have made the amendments to your config, merge it back in to Logscape. You have now upgraded your searches to deal with the new Windows App. 


OPTIION 2 - I DON'T WANT TO CHANGE MY ENVIRONMENT NOW

1. Rename your copy of the WindowsApp.bundle to NEWBUNDLEFILE
2. Rename BACKWARDSCOMPATIBILITYBUNDLE to WindowsApp.bundle
3. Zip the App up and deploy

Your environment will now be generating duplicate data: win-procinfo, win-proclist, win-procperf and win-process

It is highly advisable that you remove all dependencies in your environment as soon as possible, then use the original bundle file. 