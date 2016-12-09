/*
 * ***** BEGIN LICENSE BLOCK *****
 * Zimbra Collaboration Suite, Network Edition.
 * Copyright (C) 2005, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2016 Synacor, Inc.  All Rights Reserved.
 * ***** END LICENSE BLOCK *****
 */
/**
 * ZmCustomizeMsi.js - customize Zimbra Connector for Outlook MSI
 * Sets the default host port and connection method in the MSI
 ***/


/****************************************************
 *
 *
 *
 ****************************************************/
function PrintUsage( msg )
{
	WScript.Echo( 
	"Error: " + msg + "\n" + 
	"Usage: cscript " + WScript.ScriptName + " <msi-filename> <options>\n" +
	"    msi-filename                    : name of the Zimbra Connector for Outlook MSI\n" + 
	"    options 	\n" +
	"     Profile options 	\n" + 
	"      --profile-name,-pn  <name>|\"\"    : Zimbra profile name to create - default is Zimbra\n" + 
	"     Server options 	\n" + 
	"      --server-name,-sn <name>|\"\"      : default Zimbra server value to set in the MSI\n" + 
	"      --server-port,-sp <port>           : default port to set in the MSI\n" +
	"      --use-ssl,-ssl     <1|0>           : 1 means use HTTPS connection, 0 means HTTP connection, default is 1\n" +
	"      --proxy-setting,-prx <1|2|4>       : proxy configuration: 1 means connect directly without proxy, 2 means follow IE settings (default), 4 means manually set proxy configuration\n" +
	"      --proxy-server,-prs <name>|\"\"    : proxy host name\n" +
	"      --proxy-port,-prp <port>           : proxy server port number. default value is port 80\n" +
	"      --signature-sync,-ss <1|0>         : synchronize signatures with server, default is 1 (enabled)\n" +
	"      --share-automount,-sa <1|0>        : automatically mount shared stores, default is 1 (enabled)\n" +
	"     ZDB file options 	\n" + 
	"      --zdb-folder,-zdb <path>|\"\"      : folder to store ZDB files if roaming profiles are used. \"\" - empty path means default location\n" +
	"      --disable-autoarchive, -daa <1|0>  : 1 means that the Outlook AutoArchive process will be disabled.\n" +	
	"     Password options 	\n" + 
	"      --store-password,-pw <1|0>         : 1 means password is stored in the profile, 0 means password not stored, and Outlook will prompt for it upon entry\n" +
	"     Local rule options 	\n" + 
	"      --localrules-enabled,-lre <1|0>    : 1 means Outlook local rules enabled, 0 means Outlook local rules disabled\n" +
	"     Logging options 	\n" + 
	"      --log-enabled,-le  <1|0>           : 1 means logging enabled, 0 means logging disabled\n" +
	"      --log-filename,-lf <path>          : path to the file where log will be written. Should end with 'zco.log'. F.e. c:\\somedir\\zco.log \n" +
	"      --log-max-size,-lms <size>         : maximum size of log file in megabytes\n" +
	"     Timeout options 	\n" + 
	"      --connect-timeout,-tc <ms>         : connect timeout in milliseconds. Should be greater than 900000. if less 900000 is used\n" +
	"      --send-timeout,-ts <ms>            : send timeout in milliseconds. Should be greater than 900000. if less 900000 is used\n" +
	"      --receive-timeout,-tr <ms>         : receive timeout in milliseconds. Should be greater than 900000. if less 900000 is used\n" +
	"      --option-receive-timeout,-tor <ms> : option receive timeout in milliseconds. Should be greater than 900000. if less 900000 is used\n"+
	"     Sync failure options 	\n" + 	
	"      --inbox-failures-off,-ifo <1|0>    : 1 means create sync failure messages only in 'Sync Issues' folders, 0 means create in Sync Issues and Inbox\n" +
	"     GAL options 	\n" + 	
	"      --galsync-mode,-gsm  <2|1|0>              : 2 means GAL sync disabled, 1 means manual GAL sync only, 0 means manual and automatic GAL sync enabled\n" +
	"      --galsync-delta,-gsd  <interval>          : interval in minutes to attempt a GAL delta sync (BES only)\n" +
	"      --galsync-sleep,-gss  <interval>          : interval in milliseconds to sleep for GAL throttling.  Should be between 0 and 60000\n" +
	"      --galsync-numbeforesleep,-gsn  <number>   : number of contacts to process before sleep for GAL throttling. Should be between 2 and 500.\n" +
	"                                                  This value wouldn't be used unless galsync-sleep is specified\n" +
	"      --galsync-sort,-gso  <0|1|2>                : 0 means sort by display name, 1 means sort by fileas, 2 means use Outlook Address Book sort option\n" +
	"      --galsync-disablealiases	<1|0>		 : 1 means disable aliases in the GAL, 0 means enable aliases in the GAL\n" +
	"      --ldap-enabled,-lde  <1|0>                : 1 means LDAP enabled, 0 means LDAP disabled\n" +
	"      --ldapserver-name,-lsn <name>|\"\"        : default LDAP server name value to set in the MSI\n" +
	"     Download options 	\n" + 
	"      --download-mode,-dm  <2|1|0>       : 2 means download headers only and preserve transport headers, 1 means download headers only, 0 means download the whole message\n" +	
	"      --zdb-compact,-zc  <1|0>       : 1 means that ZCO will keep track of ZDB file size and will compact it automatically. 0 will disable automatic compacting\n" +	
    "      --disable-autoupgrade,-du <1|0>  : 1 means that ZCO will not try to automatically look for updated version on the server. 0 means that user will be prompted with offer to upgrade when updates are available\n" +
	"     ZCO Language override \n" +
	"      --language, -lang  <LCID>        : LCID must be a decimal locale ID, as per http://msdn.microsoft.com/en-gb/goglobal/bb964664.aspx (e.g. 1031 for German, default=0 i.e. no override)\n" +
	"Samples: \n"+
	"  Cscript ZmCustomizeMsi.js ZimbraOlkConnector.msi: shows current settings of MSI file \n"+
	"  Cscript ZmCustomizeMsi.js ZimbraOlkConnector.msi -sn server.zimbra.com: sets default server to be 'server.zimbra.com'\n"+
	"  Cscript ZmCustomizeMsi.js ZimbraOlkConnector.msi -sn server.zimbra.com -sp 443 -ssl 1: sets default server to be 'server.zimbra.com' and sets SSL connection method to be used\n"
	);
	
}




/****************************************************
 *
 *
 *
 ****************************************************/
function IsNumber( str )
{
	var allNums = "0123456789";
	for( var i = 0; i < str.length; i++ )
	{
		if( allNums.indexOf(str.charAt(i)) == -1 )
			return false;
	}
	return true;
}



/****************************************************
 *
 *
 *
 ****************************************************/
function MsiExists( filename )
{
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	return (fso.FileExists(filename));
}

/****************************************************
 *
 *
 *
 ****************************************************/
function MsiFolderExists( filename )
{
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	return (fso.FolderExists(filename));
}

/****************************************************
 *
 *
 *
 ****************************************************/
function UnEscapePath( filename )
{
	var Unescaped = filename.replace("^%","%");
	while(Unescaped != filename)
	{
		filename = Unescaped+"";
		Unescaped = filename.replace("^%","%");
	}
	return Unescaped;
}


/****************************************************
 *
 *
 *
 ****************************************************/
var msiOpenDatabaseModeReadOnly = 0;
var msiColumnInfoNames = 0;
var msiColumnInfoTypes = 1;
var msiDatabaseNullInteger = -2147483648;

function ShowCurrentParameters(ZimbraMsiDb)
{
	WScript.Echo("Current settings:");
	var installer = new ActiveXObject("WindowsInstaller.Installer");
	var database = installer.OpenDatabase( ZimbraMsiDb, msiOpenDatabaseModeReadOnly );
	var query = database.OpenView("select Name, Value from Registry");
	if(null != query)
	{
		query.Execute();
		
		var record = query.Fetch();
		while(null != record)
		{
			var iValue = record.IntegerData(2);
			if(msiDatabaseNullInteger == iValue)
			{
				iValue = record.StringData(2);
				if('#' == iValue.charAt(0))
				{
					//we have numeric value
					iValue = iValue.substr(1);
				}
				else
				{
					iValue = "\""+iValue+"\"";
				}
			}
			WScript.Echo("  "+record.StringData(1)+": "+iValue);
			record = query.Fetch();
		}
	}
}


/****************************************************
 *
 *
 *
 ****************************************************/
var msiOpenDatabaseModeTransact = 1;
var nArgs = WScript.Arguments.length;
if( nArgs < 1 )
{
	PrintUsage("Not enough parameters");
	WScript.Quit(1);
}

var ZmMsiDb = WScript.Arguments(0);
if( !MsiExists(ZmMsiDb) )
{
	PrintUsage("Cannot find file: " + ZmMsiDb);
	WScript.Quit(1);
}

var updates = new Array();
var updatesCount = 0;

//WScript.Sleep(30000);

//WScript.Echo("WScript.Arguments = "+WScript.Arguments.length);
var i=1;
while(i<WScript.Arguments.length)
{
	var iParameter = WScript.Arguments(i);
	//WScript.Echo(""+i+": " +iParameter);
	var iValue;
	switch(iParameter)
	{
		case "--profile-name":
		case "-pn":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			if("" == iValue)
				iValue = "Zimbra";
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ProfileName";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Property"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--server-name":
		case "-sn":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ZimbraServerName";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--server-port": 
		case "-sp": 
			iValue = WScript.Arguments(i+1);
			//WScript.Echo(iValue);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid port specified: " + iValue);
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ZimbraServerPort";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
		//WScript.Echo(""+i+": " +iParameter);
			++i;
		//WScript.Echo(""+i+": " +iParameter);
		break;
		
		case "--use-ssl":
		case "-ssl":
			iValue = WScript.Arguments(i+1);
			if( iValue != "0" && iValue != "1" )
			{
				PrintUsage("Invalid connection method." + iValue + " Specify 0 or 1.");
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ZimbraConnectionMethod";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--zdb-folder":
		case "-zdb":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			if(iValue != "" && !MsiFolderExists(iValue))
			{
   				var UnEscaped = UnEscapePath(iValue);

              	var WshShell = WScript.CreateObject("WScript.Shell");
              	var Expanded = WshShell.ExpandEnvironmentStrings(UnEscaped);
              	//we check that at least directory exists for expanded variables
              	if( !MsiFolderExists(Expanded) )
              	{
              		PrintUsage("Cannot find folder: " + UnEscaped);
              		WScript.Quit(1);
              	}
              	iValue = UnEscaped+"";
            }
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ZimbraZDBFolder";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--store-password":
		case "-pw":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "StorePassword";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		
		case "--localrules-enabled":
		case "-lre":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "EnableLocalRules";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		
		case "--log-enabled":
		case "-le":
			iValue = WScript.Arguments(i+1);
			if( iValue != "0" && iValue != "1" )
			{
				PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "LoggingEnabled";
			updates[updatesCount].val = /*"#" + */iValue;
			updates[updatesCount].tbl = "Property"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--log-filename":
		case "-lf":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			if("" == iValue)
				iValue = "?";
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "LoggingFileName";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Property"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--log-max-size":
		case "-lms":
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid parameter value: " + iValue);
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "LoggingMaxMB";
			updates[updatesCount].val = /*"#" + */iValue;
			updates[updatesCount].tbl = "Property"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--connect-timeout":
		case "-tc":
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid parameter value: " + iValue);
				WScript.Quit(1);
			}
			iValue = iValue + 0; //getting numeric value
			if(iValue < 900000)
				iValue = 900000;

			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ConnectTimeout";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		case "--send-timeout":
		case "-ts":
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid parameter value: " + iValue);
				WScript.Quit(1);
			}
			iValue = iValue + 0; //getting numeric value
			if(iValue < 900000)
				iValue = 900000;

			updates[updatesCount] = new Object();
			updates[updatesCount].col = "SendTimeout";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		case "--receive-timeout":
		case "-tr":
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid parameter value: " + iValue);
				WScript.Quit(1);
			}
			iValue = iValue + 0; //getting numeric value
			if(iValue < 900000)
				iValue = 900000;

			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ReceiveTimeout";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		case "--option-receive-timeout":
		case "-tor":
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid parameter value: " + iValue);
				WScript.Quit(1);
			}
			iValue = iValue + 0; //getting numeric value
			if(iValue < 900000)
				iValue = 900000;

			updates[updatesCount] = new Object();
			updates[updatesCount].col = "OptionReceiveTimeout";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		case "--inbox-failures-off":
		case "-ifo":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "turnOffInboxFailures";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--galsync-mode":
		case "-gsm":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1"  && iValue != "2")
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0, 1, or 2.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalSyncMode";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--galsync-delta":
		case "-gsd":
		iValue = WScript.Arguments(i+1);
		if( !IsNumber(iValue) )
		{
			PrintUsage("Invalid parameter value: " + iValue);
			WScript.Quit(1);
		}
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalDeltaSync";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--galsync-sleep":
		case "-gss":
		iValue = WScript.Arguments(i+1);
		if( !IsNumber(iValue) )
		{
			PrintUsage("Invalid parameter value: " + iValue);
			WScript.Quit(1);
		}
		if(iValue < 0)
			iValue = 0;
		else if (iValue > 60000)
			iValue = 60000;	
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalFullSyncSleep";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--galsync-numbeforesleep":
		case "-gsn":
		iValue = WScript.Arguments(i+1);
		if( !IsNumber(iValue) )
		{
			PrintUsage("Invalid parameter value: " + iValue);
			WScript.Quit(1);
		}
		if(iValue < 2)
			iValue = 2;
		else if (iValue > 500)
			iValue = 500;	
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalFullSyncSleepAfterNumContacts";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;	
		break;
		case "--galsync-sort":
		case "-gso":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" && iValue != "2" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalSort";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--galsync-disablealiases":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "GalSyncDisableAliases";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--ldap-enabled":
		case "-lde":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "EnableLDAP";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--ldapserver-name":
		case "-lsn":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "LDAPServerName";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		case "--download-mode":
		case "-dm":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1"  && iValue != "2")
		{
			PrintUsage("Invalid parameter value." + iValue + " Specify 0, 1, or 2.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "DownloadMode";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		//skip next parameter
		i++;		
		break;
		case "--zdb-compact":
		case "-zc":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value " + iValue + ". Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "ZDBAutoCompactEnabled";
		updates[updatesCount].val = "#" + iValue;
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		i++;
		break;
        case "--disable-autoupgrade":
        case "-du":
		iValue = WScript.Arguments(i+1);
		if( iValue != "0" && iValue != "1" )
		{
			PrintUsage("Invalid parameter value " + iValue + ". Specify 0 or 1.");
			WScript.Quit(1);
		}		
		updates[updatesCount] = new Object();
		updates[updatesCount].col = "SkipVersionUpgrade";
		if ( "1" == iValue) 
        {
            updates[updatesCount].val = "65535";
        } else {
            updates[updatesCount].val = "";
        }
		updates[updatesCount].tbl = "Registry"
		updatesCount++;
		i++;
		break;        
		case "--proxy-server":
		case "-prs":
			iValue = WScript.Arguments(i+1);
			iValue = iValue + "";
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ProxyServerName";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			i++;
		break;
		
		case "--proxy-port": 
		case "-prp": 
			iValue = WScript.Arguments(i+1);
			if( !IsNumber(iValue) )
			{
				PrintUsage("Invalid port specified: " + iValue);
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ProxyServerPort";
			updates[updatesCount].val = iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			++i;
		break;	
		case "--proxy-setting": 
		case "-prx": 
			iValue = WScript.Arguments(i+1);
			if( iValue != "1" && iValue != "2" && iValue != "4")
			{
				PrintUsage("Invalid option specified: " + iValue);
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "ProxyChoice";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			++i;
		break;        	
		case "--disable-autoarchive": 
		case "-daa": 
			iValue = WScript.Arguments(i+1);
			if( iValue != "1" && iValue != "0")
			{
				PrintUsage("Invalid option specified: " + iValue);
				WScript.Quit(1);
			}
			updates[updatesCount] = new Object();
			updates[updatesCount].col = "DisableAutoArchive";
			updates[updatesCount].val = "#" + iValue;
			updates[updatesCount].tbl = "Registry"
			updatesCount++;
			//skip next parameter
			++i;
		break;
        case "--signature-sync":
        case "-ss":
            iValue = WScript.Arguments(i + 1);
            if (iValue != "1" && iValue != "0") {
                PrintUsage("Invalid option specified: " + iValue);
                WScript.Quit(1);
            }
            updates[updatesCount] = new Object();
            updates[updatesCount].col = "SigSyncEnabled";
            updates[updatesCount].val = "#" + iValue;
            updates[updatesCount].tbl = "Registry"
            updatesCount++;
            //skip next parameter
            ++i;
            break;
        case "--share-automount":
        case "-sa":
            iValue = WScript.Arguments(i + 1);
            if (iValue != "1" && iValue != "0") {
                PrintUsage("Invalid option specified: " + iValue);
                WScript.Quit(1);
            }
            updates[updatesCount] = new Object();
            updates[updatesCount].col = "ShareAutomountEnabled";
            updates[updatesCount].val = "#" + iValue;
            updates[updatesCount].tbl = "Registry"
            updatesCount++;
            //skip next parameter
            ++i;
            break;
        case "--language":
        case "-lang":
            iValue = WScript.Arguments(i + 1);
            updates[updatesCount] = new Object();
            updates[updatesCount].col = "ZCOLanguage";
            updates[updatesCount].val = "#" + iValue;
            updates[updatesCount].tbl = "Registry"
            updatesCount++;
            //skip next parameter
            ++i;
            break;
        default:
			PrintUsage("Invalid parameter: " + iParameter);
	}	
	//WScript.Echo(""+i+": " +iParameter);
	i++;
}


if(0 == updatesCount)
{
	ShowCurrentParameters(ZmMsiDb);
	WScript.Quit(0);
}

/*
var ZmServer = WScript.Arguments(1);
var ZmPort = WScript.Arguments(2);
var ZmUseSSL = WScript.Arguments(3);
var ZDBFolder = WScript.Arguments(4);

if( ZDBFolder != "" && !MsiFolderExists(ZDBFolder) )
{
	var UnEscaped = UnEscapePath(ZDBFolder);

	var WshShell = WScript.CreateObject("WScript.Shell");
	var Expanded = WshShell.ExpandEnvironmentStrings(UnEscaped);
	//we check that at least directory exists for expanded variables
	if( !MsiFolderExists(Expanded) )
	{
		PrintUsage("Cannot find folder: " + UnEscaped);
		WScript.Quit(1);
	}
	ZDBFolder = UnEscaped+"";
}


if( !IsNumber(ZmPort) )
{
	PrintUsage("Invalid port specified");
	WScript.Quit(1);
}

if( ZmUseSSL != "0" && ZmUseSSL != "1"  )
{
	PrintUsage("Invalid connection method.  Specify 0 or 1.");
	WScript.Quit(1);
}

var updates = new Array(3);
updates[0] = new Object();
updates[0].col = "ServerName";
updates[0].val = ZmServer;
updates[1] = new Object();
updates[1].col = "ServerPort";
updates[1].val = "#" + ZmPort;
updates[2] = new Object();
updates[2].col = "ConnectionMethod";
updates[2].val = "#" + ZmUseSSL;
updates[3] = new Object();
updates[3].col = "ZDBFolder";
updates[3].val = ZDBFolder;
*/


WScript.Echo("Setting parameters: \n");
for(var i = 0; i< updates.length; i++)
{
	var iValue = updates[i].val;
	if('#' == iValue.charAt(0))
		iValue = iValue.substr(1);
	else
		iValue = "\""+iValue+"\"";
	WScript.Echo(updates[i].col+": "+iValue);
}

//for debug only
//WScript.Quit(0);

try
{
	var installer = new ActiveXObject("WindowsInstaller.Installer");
	var database = installer.OpenDatabase( ZmMsiDb, msiOpenDatabaseModeTransact );
	
	for( var i = 0; i < updates.length; i++ )
	{
	    var query;
	    if (updates[i].tbl == "Registry") {
	        query = "update Registry Set Value='" + updates[i].val + "' " +
    		        "Where Name='" + updates[i].col + "'";
	    }
	    else if (updates[i].tbl == "Property") {
	        query = "update Property Set Value='" + updates[i].val + "' " +
	    	        "Where Property='" + updates[i].col + "'";
	    }
	    else {
	        // Should never get here
	        WScript.Echo("Unrecognised update for '" + updates[i].tbl + "'");		
	    }
//for debug only
//		WScript.Echo("Updating database query: '"+query+"'");		
		var view = database.OpenView(query);
		view.Execute();
	}
	
	database.Commit();
	WScript.Echo("Installer db updated successfully");
}
catch(e)
{
	WScript.Echo("Error opening or updating MSI" );
	WScript.Echo("Error Number     : " + (e.number & 0xFFFF) );
	WScript.Echo("Error Description: " + e.description);
}

WScript.Quit(0);
