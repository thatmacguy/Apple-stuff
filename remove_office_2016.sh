#!/bin/sh

#Get Currently Loggedin User
loggedinuser=`python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");'`
echo "Current user is " $loggedinuser

#Halt All MS Apps
killall "Microsoft Excel"
killall "Microsoft Word"
killall "Microsoft OneNote"
killall "Microsoft Powerpoint"
killall "Microsoft Outlook"
killall "Microsoft Database Daemon"
killall "Microsoft AU Daemon"
echo "Processes stopped"

#Remove Apps
rm -rf /Applications/Microsoft\ Excel.app
rm -rf /Applications/Microsoft\ Word.app
rm -rf /Applications/Microsoft\ OneNote.app
rm -rf /Applications/Microsoft\ Powerpoint.app
rm -rf /Applications/Microsoft\ Outlook.app
echo "apps removed"

#Remove Licensing Files
rm -rf /Library/LaunchDaemons/com.microsoft.office.licensingV2.helper.plist
rm -rf /Library/PrivilegedHelperTools/com.microsoft.office.licensingV2.helper
rm -rf /Library/Preferences/com.microsoft.office.licensingV2.plist
echo "Licensing Files removed"

#Remove Current User Library Files
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.errorreporting
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.netlib.shipassertprocess
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.Office365ServiceV2
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.Outlook
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.Powerpoint
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.RMS-XPCService
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.Word
rm -rf /User/$loggedinuser/Library/Containers/com.microsoft.onenote.mac
echo "User Library Files removed"

#Backup Outlook Data
mkdir -p /tmp/$loggedinuser/outlook2016bkup
mv /User/$loggedinuser/Library/Group\ Containers/UBF8T346G9.MS /tmp/$loggedinuser/outlook2016bkup/UBF8T346G9.MS
mv /User/$loggedinuser/Library/Group\ Containers/UBF8T345G9.Office /tmp/$loggedinuser/outlook2016bkup/UBF8T345G9.Office
mv /User/$loggedinuser/Library/Group\ Containers/UBF8T346G9.OfficeOsfWebHost /tmp/$loggedinuser/outlook2016bkup/UBF8T346G9.OfficeOsfWebHost
echo "Outlook Data moved to /tmp"

exit 0
