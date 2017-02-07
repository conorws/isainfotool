/*
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    This code is Copyright (c) 2004 - 2006 Microsoft Corporation.  

    All rights reserved.

    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
    PARTICULAR PURPOSE.

    IN NO EVENT SHALL MICROSOFT AND/OR ITS RESPECTIVE SUPPLIERS BE
    LIABLE FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
    DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
    WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
    ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
    OF THIS CODE OR INFORMATION.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

   Purpose: Gathers ISA Server 2004 information and saves it to XML
     
   Requirements: 
         - ISA 2004 Server or ISA 2004 Admin objects on the local host
         - Access rights to local or remote ISA 2004 server for interactive
           account
  
   Version:
         1.0           09/17/2004 - RTM
         1.0.2161.2    11/07/2004 
                            - cleaned up the share ACL enumeration
                            - fixed export and xml save error handling
                            - fixed share enumeration error handling
                            - added Enterprise export
                            - added cmd-line options
                            - updated GetISA to support future merged script
                            - added debug output
         1.0.2161.3    11/08/2004
                            - fixed xpath statements for array & server for EE
                            - made extra logging properties SE-selective
         1.0.2161.4    12/01/2004
                            - fixed server enumerations & function results
                            - added some file saves during scanning
         1.0.2161.5    12/05/2004
                             - added "/serveronly" flag for EE extra servers
                             - added "/quiet" flag to block error popups
         1.0.2161.6    01/17/2005
                             - fixed local server name comparison
         1.0.2161.7    01/31/2005
                             - fixed server scanning if export method fails
         1.0.2161.8    06/18/2005
                             - updated event log scanning
                             - cleaned up ToHex function
         1.0.2161.9    06/29/2005
                             - added IIS enumeration
         1.0.2161.10 08/01/2005
                             - added Winsock catalog dump for Win2K3
                             - added DHCP Server dump
         1.0.2161.11 08/22/2005
         					 - Updated to support IsaBPA tool
         1.0.2161.12 10/23/2005
                             - Fixed "invalid character" errors in XML.
         1.0.2161.13 10/25/2005
                             - Fixed ConnectToCss bug.
         1.0.2161.14 10/27/2005
                             - Fixed SE bug caused by ConnectToCss fix.
         1.0.2161.15 12/16/2005
                             - Fixed "ipconfig /all" cmd
         1.0.2161.16 01/03/2006
                             - Fixed DBCS issue in WMI files query
         1.0.2161.17 01/23/2006
                             - Fixed character escaping & dropped password
         1.0.2161.18 02/17/2006
                             - re-fixed character escaping
         1.0.2161.19 02/23/2006
                             - added version tracking for .hta
         1.0.2161.20 06/08/2006
                             - fixed XML reformatting & replaced character 
                               escaping with CDATA nodes
                             - updated external command process
                             - added netsh ipsec sta sho all to GetExtData()
         1.0.2161.21 06/21/2006
                             - fixed xml cleanup bug
                             - fixed WScript.Shell.Exec() method hanging
         1.0.2161.22 05/17/2007
                             - fixed /array flag and made it default
                             - added CSS connection options (EE)
                             - added remote array connection options (SE)
                             - limited event log WMI query to 100 entries
                             - fixed 0-length output in special cases
                             - added uninstall registry data
                             - fixed remote server message runtime error
         1.0.2161.23 07/12/2007
                             - added /enterprise option
                             - updated helptext
         1.0.2161.24 02/11/2007
                             - added ability to read local (lm)hosts file 
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
*/

/*
 * JScript can"t pass most variables "byref", so we construct custom objects to 
 * hold global objects & variables
 */
var g_oObjects     = new IsaInfoObjects();
var g_oVariables    = new IsaInfoVariables();
var g_oMessages    = new IsaInfoMessages();


Main();

/**********************************************************************
 * Main()
 * This function:
 *    1. validates that the script is running under "cscript.exe" and 
 *        restarts the script if not (avoids popups from console output)
 *    2. calls into
 *        InitEnvironment()
 *      ExportIsaXml()
 *      GetIsa2K4ServerInfo()
 *        CloseFiles()
 *  3. called by 
 *        user
 *
 * if successful:
 *    1. returns g_oVariables.lS_OK
 *    2. ISA XML configuration contains additional information inserted 
 *        into the relevant Server node as child elements
 *
 * if unsuccessful:
 *    1. returns integer representing the general failure location
 *    2. ServersNode contents are not guaranteed; depends on the failure point
 *********************************************************************/
function Main()
{
    EnterFunction( arguments );

    var iObjFailed = 1;
    var iInitFailed = 2;
    var iAdminOnly = 3;
    var iScanFailed = 4;
    var iCleanupFailed = 5;
    var iArgsFailed = 7;
    var iHidden = 10;
    
    if( GetObjects() != g_oVariables.lS_OK ) 
    {
        WScript.Quit( iObjFailed );
    }

    var szScriptEngine = WScript.FullName.toLowerCase();
    var iStart = szScriptEngine.lastIndexOf( "\\" ) + 1;
    var iEnd = szScriptEngine.lastIndexOf( "." ) - iStart;

    g_oMessages.szScriptEngine = szScriptEngine.substr( iStart, iEnd );

    if( ParseArgs() != true )
    {
        ShowUsage();
        WScript.Quit( iArgsFailed );
    }

    if( g_oVariables.fDebugMode )
    {
        LogMessage( "    --> Main()" );
    }

    /*
     * avoids Wscript popups where we want cmd line output, since
     * "wscript.exe" is the default scripting engine
     */
    if ( g_oMessages.szScriptEngine != "cscript" )
    {
        WScript.Quit( g_oObjects.oWsh.run( "cscript \"" + 
                            WScript.ScriptFullName + "\"" + 
                            g_oMessages.szCmdOpts, 
                            iHidden, true ) );
    }

    if ( InitEnvironment() != g_oVariables.lS_OK )
    {
        WScript.Echo( g_oMessages.L_SetupFailed_txt );
        CloseFiles();
        WScript.Quit( iInitFailed );
    }

    ExportIsaXml();
    var iRtn = CloseFiles();
    ExitFunction( arguments, iRtn );
    WScript.Quit( iRtn );
}

/**********************************************************************
 * InitEnvironment( )
 * This function:
 *    1. populates various global variables
 *    2. calls into
 *        LogMessage()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns integer representing the general failure point or err
 *        from call to other function
 *    2. ServersNode contents are dependent on the failure point
 *********************************************************************/
function InitEnvironment( )
{
    EnterFunction( arguments );

    var GetIsaFailed = 1;
    var szMpsReports = "";
    
    g_oVariables.szComSpec = g_oObjects.oWsh.ExpandEnvironmentStrings( "%ComSpec%" );
    g_oVariables.szSysFolder = g_oObjects.oWsh.ExpandEnvironmentStrings( "%WinDir%" ) + 
                                "\\system32\\";
    g_oVariables.szThisServer = 
            g_oObjects.oWsh.ExpandEnvironmentStrings( "%ComputerName%" ).toLowerCase();
    g_oVariables.szThisUser = 
            g_oObjects.oWsh.ExpandEnvironmentStrings( "%UserDomain%" ) +
                "\\" + 
                g_oObjects.oWsh.ExpandEnvironmentStrings( "%UserName%" );
    if( !g_oVariables.fOneServerSet )
    {
        g_oMessages.szThisServer = g_oVariables.szThisServer;
    }

    if( !g_oVariables.fPathSet )
    {
        g_oVariables.szISAInfoPath = 
                g_oObjects.oWsh.SpecialFolders( "Desktop" ) + "\\";
    }

    g_oMessages.szHeaderMsg = "\r\n" + g_oMessages.szDivider + "\r\n" + 
                g_oMessages.L_TitleMsg_txt + 
                "\r\n" + g_oMessages.L_RunningOn_txt + 
                g_oVariables.szThisServer + g_oMessages.L_RunningAs_txt + 
                g_oVariables.szThisUser + "\r\n" + g_oMessages.L_Start_txt + 
                new Date().toString() + "\r\nas \"" + g_oMessages.szScriptName +
                g_oMessages.szCmdOpts + "\"\r\n\r\n" + g_oMessages.szDivider;

    /*
     * If ISAInfo is called by MPSReports, the %MPSReports% 
     * environment variable contains the full path to the 
     * output destination path for all utilities.
     * If the env var doesn"t exist, the ExpandEnvironmentStrings 
     * method returns the string passed into it
     */
    szMpsReports = g_oObjects.oWsh.ExpandEnvironmentStrings( "%MPSReports%\\" );
    if ( szMpsReports != "%MPSReports%\\" )
    {
        g_oVariables.fMPSReports = true;
        g_oVariables.szISAInfoPath = szMpsReports;
    }

    g_oVariables.szXmlFile = g_oVariables.szISAInfoPath + 
                            g_oMessages.szScriptName + "_" + 
                                g_oVariables.szThisServer + ".xml";

    g_oVariables.szTempFilePath = g_oVariables.szISAInfoPath + 
                                g_oMessages.szScriptName + 
                                "_tempdata.txt";
    if( StartLog() != g_oVariables.lS_OK )
    {
        ExitFunction( arguments, -1 );
        return -1;
    }

    if ( !DetermineIsaEnvironment() )
    {
        ExitFunction( arguments, GetIsaFailed );
        return GetIsaFailed;
    }
    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
}

/**********************************************************************
 * GetObjects( )
 * This function:
 *    1. populates various global objects
 *    2. calls into
 *        LogError()
 *        ObjFactory()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns integer representing the general failure point or err
 *        from call to other function
 *    2. ServersNode contents are dependent on the failure point
 *********************************************************************/
function GetObjects()
{
    EnterFunction( arguments );

    try
    {
        g_oObjects.oWsh = ObjFactory( "WScript.Shell" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_NoWsh_txt );
        ExitFunction( arguments, err.number );
        return err.number;
    }
    
    try
    {
        g_oObjects.oIsaXml = ObjFactory( "MSXML2.DomDocument.3.0" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_NoIsaData_txt );
        ExitFunction( arguments, err.number );
        return err.number;
    }
    g_oObjects.oIsaXml.async = false;

    try
    {
        g_oObjects.oFSO = ObjFactory( "Scripting.FileSystemObject" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_NoFso_txt );
        ExitFunction( arguments, err.number );
        return err.number;
    }

    ExitFunction( g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
}

/*@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
 # Start of hta C&P section
 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@*/
/*#######################################
 # server data gathering tools
 # most depend on a valid connection to 
 # WMI on the target server
 ######################################*/
/**********************************************************************
 * GetServerData( oParentNode )
 * This function:
 *    1. Derives the server name from the XMLDomNode data
 *        obtains additional Server data in the proper order
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4ServerData()
 *        GetWmiObjects()
 *        GetOsData()
 *        GetHardwareData()
 *        GetEvtLogData()
 *        GetExtData()
 *        ReadFileContents()
 *        GetIsaFilesData()
 *        GetRegistryData()
 *        CheckIIS()
 *  3. called by 
 *        GetIsa2K4ServerInfo( )
 *
 * if successful:
 *    1. Each fpc4:Server node contains additional data rooted in 
 *        IsaServer child node
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to GetServerData()
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
 function GetServerData( oParentNode )
 {
    EnterFunction( arguments );

    var oNewNode;                //general xmldomelement
    var szServerName = g_oObjects.oIsaServer.Name.toLowerCase();
    var iRtn = 0;

    if( g_oVariables.szThisServer != szServerName || g_oVariables.szCss )
    {
        LogMessage( Fprintf( g_oMessages.L_SvrNotLocal_txt, new Array( szServerName ) ) );
        ExitFunction( arguments, iRtn );
        return iRtn;
    }

    if ( !GetWmiObjects( szServerName ) )
    {
        LogMessage( g_oMessages.L_SvrConnErr_txt + szServerName );
        ExitFunction( arguments, iRtn );
        return iRtn;
    }

    GetOSData( oParentNode );
    GetHardwareData( oParentNode );
    GetEvtLogData( oParentNode );

    if ( szServerName == g_oVariables.szThisServer )
    {
        g_oVariables.fIsW2K3 = ( oParentNode.selectSingleNode( 
                                    "//BuildNumber" ).text == "3790" );
        g_oVariables.fLocal_RRAS = IsLocalRras( );
    }

    GetShares( oParentNode );

    /*
     * because MPSReports needs a quick run
     */
    if ( !g_oVariables.fMPSReports )
    {
        oNewNode = NewChildNode( oParentNode, "Files", "" );
        GetIsaFilesData( oNewNode, 
                oParentNode.selectSingleNode( "InstallationDirectory" ).text,
                szServerName );

        oNewNode = NewChildNode( oParentNode, "Registry", "" );
        GetRegistryData( oNewNode )
    }

    CheckIIS( oParentNode, szServerName );
    oNewNode = NewChildNode( oParentNode, "NetInfo", "" );
    GetExtData( oNewNode );

    ExitFunction( arguments, iRtn );
    return iRtn;

}

/**********************************************************************
 * GetOSData( oParentNode )
 * This function:
 *    1. populates oParentNode with OS-related WMI properties
 *    2. calls into
 *        GetWmiObjectsInfo()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. creates new IsaServer childnodes as:
 *        Win32_OperatingSystem
 *        Win32_PageFile
 *        Win32_QuickFixEngineering
 *        Win32_SystemDriver
 *        Win32_Service
 *        ..each containing data based on the query defined
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to GetWmiObjectsInfo()
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
 function GetOSData( oParentNode )
 {
    EnterFunction( arguments );

     var szWBemClass;
     var oOsNode;
    var oNewNode;
    var szOsQuery      = "select BuildNumber, BuildType, Caption, CSDVersion, " +
                        "OSLanguage, OSProductSuite, SerialNumber, " +
                        "ServicePackMajorVersion, ServicePackMinorVersion, " +
                        "SuiteMask, SystemDevice, SystemDirectory, " +
                        "TotalVirtualMemorySize, Version, WindowsDirectory " +
                        "from Win32_OperatingSystem";
    var szPfQuery      = "select Compressed, FileSize, FSName, InitialSize, " +
                        "MaximumSize, Name from Win32_PageFile";
    var szHfQuery      = "select Description, HotFixID from Win32_QuickFixEngineering " +
                        "where HotFixID != \"File 1\"";
    var szDrvQuery      = "select DisplayName, Description, Name, PathName, Started, " +
                        "StartMode, StartName, State from Win32_SystemDriver";
    var szSvcQuery      = "select DisplayName, Description, Name, PathName, " +
                        "ProcessId, StartMode, StartName, State from Win32_Service";
    var szProcQuery  = "Select CreationDate, ExecutablePath, HandleCount, " +
                        "KernelModeTime, Name, ParentProcessId, Priority, " +
                        "ProcessId, UserModeTime from Win32_Process";

    try
    {
        szWBemClass = "Win32_OperatingSystem";
        GetWmiObjectsInfo( oParentNode, szOsQuery, "Win32_OperatingSystem", "OsInfo" );
        oOsNode = oParentNode.selectSingleNode( "Win32_OperatingSystem" );
        szWBemClass = "Win32_PageFile";
        GetWmiObjectsInfo( oOsNode, szPfQuery, "Win32_PageFile", "PageFile" );
        szWBemClass = "Win32_QuickFixEngineering";
        GetWmiObjectsInfo( oOsNode, szHfQuery, "Win32_QuickFixEngineering", "Hotfix" );
        szWBemClass = "Win32_SystemDriver";
         GetWmiObjectsInfo( oOsNode, szDrvQuery, "Win32_SystemDriver", "Driver" );
        szWBemClass = "Win32_Service";
        GetWmiObjectsInfo( oOsNode, szSvcQuery, "Win32_Service", "Service" );
        szWBemClass = "Win32_Process";
        GetWmiObjectsInfo( oOsNode, szProcQuery, "Win32_Process", "Process" );

        oNewNode = NewChildNode( oOsNode, "Boot.ini", "" );
        ReadFileContents( oNewNode, "c:\\boot.ini" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_QueryFailed_txt + szWBemClass );
    }

    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
 }

/**********************************************************************
 * GetHardwareData( oParentNode )
 * This function:
 *    1. populates oParentNode with computer hardware-related WMI properties
 *    2. calls into
 *        GetWmiObjectsInfo()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. creates new IsaServer childnodes as:
 *        Win32_ComputerSystem
 *        Win32_Processor
 *        Win32_DiskDrive
 *        Win32_POTSModem
 *        Win32_NetworkAdapter
 *        ..each containing data based on the query defined
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to GetWmiObjectsInfo()
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
 function GetHardwareData( oParentNode )
 {
    EnterFunction( arguments );

     var szWBemClass;
    var szHwQuery    = "select BootupState, CurrentTimeZone, Description, Domain, " +
                        "DomainRole, Manufacturer, Model, NumberOfProcessors, " +
                        "TotalPhysicalMemory from Win32_ComputerSystem";
    var szCpuQuery    = "select Availability, CPUStatus, CurrentClockSpeed, DeviceID, " +
                        "l2CacheSize, l2CacheSpeed, Name from Win32_Processor";
    var szDiskQuery    = "select Caption, Description, DeviceID, DriveType, FileSystem, " +
                        "FreeSpace, LastErrorCode, Size, VolumeDirty, VolumeName from " +
                        "Win32_LogicalDisk";
    var szDriveQuery = "select Availability, CapabilityDescriptions, Caption, " +
                        "Description, DeviceID, InterfaceType, LastErrorCode, " +
                        "Manufacturer, Model, Name, Partitions, SCSIBus, " +
                         "SCSILogicalUnit, SCSIPort, SCSITargetId, Size, Status from " +
                        "Win32_DiskDrive";
    var szMdmQuery    = "select AnswerMode, AttachedTo, Availability, Caption, " +
                        "ConfigurationDialog, Description, MaxBaudRateToPhone, " +
                        "MaxBaudRateToSerialPort, Model, ModemInfPath, Status " +
                        "from Win32_POTSModem";
    var szNicQuery    = "select AdapterType, Availability, ConfigManagerErrorCode, " +
                        "Description, DeviceID, Index, LastErrorCode, Name, " +
                        "MACAddress, MaxNumberControlled, MaxSpeed, NetConnectionStatus, " +
                        "PowerManagementCapabilities, ServiceName, Speed, StatusInfo, " +
                        "TimeOfLastReset from Win32_NetworkAdapter";

    try
    {
        szWBemClass = "Win32_ComputerSystem"
        GetWmiObjectsInfo( oParentNode, szHwQuery, "Hardware", szWBemClass );

        szWBemClass = "Win32_DiskDrive";
        GetWmiObjectsInfo( oParentNode.selectSingleNode( "Hardware" ), szDriveQuery, 
                            szWBemClass, "Disk" );

        szWBemClass = "Win32_LogicalDisk";
        GetWmiObjectsInfo( oParentNode.selectSingleNode( "Hardware" ), szDiskQuery, 
                            szWBemClass, "Drive" );

        szWBemClass = "Win32_NetworkAdapter";
        GetWmiObjectsInfo( oParentNode.selectSingleNode( "Hardware" ), szNicQuery, 
                            szWBemClass, "NIC" );

        szWBemClass = "Win32_Processor";
        GetWmiObjectsInfo( oParentNode.selectSingleNode( "Hardware" ), szCpuQuery, 
                            szWBemClass, "CPU" );

        szWBemClass = "Win32_POTSModem";
        GetWmiObjectsInfo( oParentNode.selectSingleNode( "Hardware" ), szMdmQuery, 
                            szWBemClass, "Modem" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_QueryFailed_txt + szWBemClass );
    }

    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
 }

/**********************************************************************
 * GetRegistryData( oParentNode )
 * This function:
 *    1. populates oParentNode with registry data from selected keys
 *    2. calls into
 *        EnumRegKeys()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. creates new IsaServer childnodes as:
 *        HKLM
 *            Key...
 *        ..each containing data from the specified registry tree
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to EnumRegistryKeys()
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
 function GetRegistryData( oParentNode )
 {
    EnterFunction( arguments );

    var oNewNode = null;
    var iInx;
    var arrHklmKeys = new Array( "SOFTWARE\\Microsoft\\fpc",
                    "SOFTWARE\\Microsoft\\Microsoft SQL Server",
                    "SOFTWARE\\Microsoft\\MSSQLServer",
                    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
                    "SYSTEM\\CurrentControlSet\\Services\\bwcperf", 
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\ACECLIENT", 
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\BwcPerf", 
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\DNS intrusion detection filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\FTP Access Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\ISA Server H.323 Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\ISA Server RPC Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\ISA Server Streaming Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft Firewall",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server Control",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server Job Scheduler",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server ReportGenerator",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server Storage",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server Web Proxy",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\Microsoft ISA Server WMS Proxy",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\POP Intrusion Detection Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\SmtpEvt",
                    "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\SOCKS Filter",
                    "SYSTEM\\CurrentControlSet\\Services\\Fweng",
                    "SYSTEM\\CurrentControlSet\\Services\\Fwsrv",
                    "SYSTEM\\CurrentControlSet\\Services\\H323Fltr",
                    "SYSTEM\\CurrentControlSet\\Services\\isactrl",
                    "SYSTEM\\CurrentControlSet\\Services\\isasched",
                    "SYSTEM\\CurrentControlSet\\Services\\isastg",
                    "SYSTEM\\CurrentControlSet\\Services\\MSSQL$MSFW",
                    "SYSTEM\\CurrentControlSet\\Services\\NetBT",
                    "SYSTEM\\CurrentControlSet\\Services\\RemoteAccess",
                    "SYSTEM\\CurrentControlSet\\Services\\SocksFlt",
                    "SYSTEM\\CurrentControlSet\\Services\\tcpip",
                    "SYSTEM\\CurrentControlSet\\Services\\W3PCache",
                    "SYSTEM\\CurrentControlSet\\Services\\W3Proxy"
                    );

    oNewNode = NewChildNode( oParentNode, "HKLM", "" );
    for ( iInx in arrHklmKeys )
    {
        try
        {
            EnumRegistryKeys( oNewNode, g_oVariables.lHKLM, arrHklmKeys[ iInx ] );
        }
        catch( err )
        {
            LogError( err, g_oMessages.L_ReadFailed_txt + arrHklmKeys[ iInx ] + "." );
        }
        SaveXML( g_oObjects.oIsaXml );
        WScript.StdOut.WriteLine();
    }
    SaveXML( g_oObjects.oIsaXml );
    WScript.StdOut.WriteLine();
    
    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
 }

/**********************************************************************
 * GetEvtLogData( oParentNode )
 * This function:
 *    1. populates oParentNode with event log data from the Application,
 *        System and Security logs for the past 7 days
 *    2. calls into
 *        GetWmiObjectsInfo()
 *        LogMessage()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. creates new IsaServer childnodes containing event log data
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to EnumRegistryKeys()
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
function GetEvtLogData( oParentNode )
{
    EnterFunction( arguments );

    var oLogsNode;
    var oNewNode;                //general xmldomelement
    var oEventLogs;            //xmldomelement used for event log data
    var iInx;                    //array indexing var
    var DaysDiff    = -2;        //number of days to search back
    var WmiDate;
    var arrQuery    = new Array( "Application", "Security", "System" );
    var Query1        = "select ComputerName, Data, EventCode, EventIdentifier, "+
                        "EventType, InsertionStrings, Message, SourceName, " +
                        "RecordNumber, TimeGenerated, TimeWritten from " +
                        "Win32_NTLogEvent where LogFile = \"";
    /*
     * this limits the query to warning, error and audit failure events
     */
    var Query2        = "\" and ( EventType = \"1\" or EventType = \"2\" or " +
                        "EventType = \"5\" ) and TimeGenerated > \"";

    /*
     * using WMI-formatted date speeds up the query immensely
     */
    WmiDate = GetWmiDate( DaysDiff );
    LogMessage( g_oMessages.L_GetEvtLogs_txt );

    oLogsNode = NewChildNode( oParentNode, "EventLogs", "" );
    for ( iInx in arrQuery )
    {
        oNewNode = NewChildNode( oLogsNode, arrQuery[ iInx ], "" );
        if ( GetWmiObjectsInfo( oNewNode, Query1 + arrQuery[ iInx ] + 
                Query2 + WmiDate + "\"", "", "LogEntry" ) != g_oVariables.lS_OK )
        {
            return true;
        }
    }
    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
}
/*#######################################
 # End of server data gathering tools
 ######################################*/

/*#######################################
 # external utility functions
 # NOTE: these only work for the local machine, 
 # so this data will not be present for 
 # remote servers (EE)
 ######################################*/
/**********************************************************************
 * GetExtData( oParentNode )
 * This function:
 *    1. populates oParentNode with data obtained from external utilities
 *    2. calls into
 *        DoSystemCmd()
 *        LogMessage()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. creates new IsaServer childnodes containing data as specifed in
 *        the command line arrays
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
function GetExtData( oParentNode )
{
    EnterFunction( arguments );

    var arrCmdsNet    = new Array( "ipconfig /all", 
                                    "netstat -r",    
                                    "netsh show helper",
                                    "netsh ipsec sta sho all",
                                    "netsh ipsec dyn sho all",
                                    "wlbs display",
                                    "netsh dhcp server dump"
                                    );

    var arrCmdsRRAS = new Array( "netsh routing ip show interface",
                                    "netsh routing ip show persistentroutes",
                                    "netsh routing ip show rtmdestinations",
                                    "netsh routing ip show rtmroutes",
                                    "netsh routing ip relay show global",
                                    "netsh routing ip relay show ifconfig"
                                    );

    var arrNodeNames = new Array( "IPConfig", "RoutingTable", "NetshHlpr", "IpSecSta", "IpSecDyn",
                                    "WlbsDisplay", "Dhcp", "RrasIntfc", "RrasPersRoutes", 
                                    "RrasRtmDest", "RrasRtmRoutes", "RrasRlyGlobal", 
                                    "RrasRlyIfCfg"
                                    );

    var iCounter1;
    var iCounter2;
    var vItem;
    
    //read some error-producing data
    oNewNode = NewChildNode( oParentNode, "Hosts", "" );
    var szHostsPath = g_oVariables.szSysFolder + "drivers\\etc\\hosts";
    ReadFileContents( oNewNode, szHostsPath, true );
    oNewNode = NewChildNode( oParentNode, "LmHosts", "" );
    szHostsPath = g_oVariables.szSysFolder + "drivers\\etc\\lmhosts";
    ReadFileContents( oNewNode, szHostsPath, true );

    for ( iCounter1 in arrCmdsNet )
    {
        var szMsg = DoSystemCmd( arrCmdsNet[ iCounter1 ] );
        NewChildNode( oParentNode, arrNodeNames[ iCounter1 ], szMsg );
    }
    
    if ( g_oVariables.fLocal_RRAS )
    {
        for ( iCounter2 in arrCmdsRRAS )
        {
            iCounter1++;
            var szMsg = DoSystemCmd( arrCmdsRRAS[ iCounter2 ] );
            NewChildNode( oParentNode, arrNodeNames[ iCounter1 ], szMsg );
        }
    }

    /*
     * WinXP and later added the -o option to netstat to show the 
     * PID-to-binding association
     */
    var szCmd = "NETSTAT -AN";
    if ( g_oVariables.fIsW2K3 )
    {
        szCmd += "O";
    }
    var szMsg = DoSystemCmd( szCmd );
    NewChildNode( oParentNode, "Netstat", szMsg );

    /*
     * WinXP and later added a new netsh command to enable Winsock
     * catalog evaluation
     */
    if ( g_oVariables.fIsW2K3 )
    {
        szMsg = DoSystemCmd( "netsh winsock show catalog full" );
        NewChildNode( oParentNode, "Winsock", szMsg );
    }

    ExitFunction( arguments, null );
    WScript.StdOut.Write( ".\r\n" );
}

/**********************************************************************
 * DoSystemCmd( szCmd )
 * This function:
 *    1. calls the external utility with the command line passed in from
 *        GetExtData()
 *    2. calls into
 *        LogMessage()
 *        ObjFactory()
 *  3. called by 
 *        GetExtData()
 *
 * if successful:
 *    1. creates a new IsaServer childnode as specified in szCmd
 *
 * if unsuccessful:
 *    1. returns status from WshShell.Exec
 *    2. IsaServer child node contents are dependent on the failure point
 *********************************************************************/
function DoSystemCmd( szCmd )
{
    EnterFunction( arguments );

    szCmd = Fprintf( "%1cmd.exe /c %1%2", new Array( g_oVariables.szSysFolder, szCmd ) );
//    szCmd = g_oVariables.szSysFolder + szCmd;
    LogMessage( g_oMessages.L_RunCmd_txt + szCmd );

    var oJob = null;
    
    try
    {
        oJob = g_oObjects.oWsh.Exec( szCmd );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_RunFailed_txt + szCmd );
        ExitFunction( arguments, err.description );
        return err.description;
    }

    var iMax = 40;          //timeout factor of 10 secs
    var szMsg = "";
    var szErr = "";
    while( iMax  && !oJob.Status )
    {
        szMsg += oJob.StdOut.ReadAll();
        szErr += oJob.StdErr.ReadAll();
        WScript.Sleep( 250 ); //snooze for .25 second
        iMax--;
    }

    szMsg = szErr + "\r\n\r\n" + szMsg;
    if( !iMax )
    {
        LogMessage( "oJob.Status == " + oJob.Status.toString() );
        szMsg = "Timeout expired waiting for " + szCmd + " to exit.\r\n\r\n" + szMsg;
    }

    SaveXML( g_oObjects.oIsaXml );
    ExitFunction( arguments, szMsg );
    return szMsg;

}

/*#######################################
 # end of external utility functions
 ######################################*/

/*#######################################
 # file support functions
 ######################################*/
/**********************************************************************
 * StartLog()
 * This function:
 *    1. tries to create or open the log file
 *    2. calls into
 *        LogMessage()
 *        ObjFactory()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. all temp, log and XML files are Opend out properly
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns status from either WshShell.run or ReadFileContents
 *********************************************************************/
function StartLog()
{
    EnterFunction( arguments );

    var iForAppending = 8;
    var iTristateUseDefault = -2;
    var szTraceFile = "";
    var szTempFileName = "";
    var szTraceFile = g_oVariables.szISAInfoPath + 
                        g_oMessages.szScriptName + "_" +
                        g_oVariables.szThisServer + ".log";

    try
    {
        g_oObjects.fsTraceFile = g_oObjects.oFSO.OpenTextFile( szTraceFile, 
                                iForAppending, true, iTristateUseDefault );
    }
    catch( err )
    {
    //since file logging fails, we need to leave a record in the event log
        g_oObjects.oWsh.LogEvent( 1, g_oMessages.L_NoLogFile_txt + 
                g_oVariables.szTraceFile + 
                g_oMessages.L_ErrNum_txt + ToHex( err.number ) + 
                g_oMessages.L_ErrDesc_txt + err.description );
        ExitFunction( arguments, ToHex( err.number ) );
        return err.number;
    }

    LogMessage( g_oMessages.szHeaderMsg );
    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
}


/**********************************************************************
 * CloseFiles()
 * This function:
 *    1. tries to close all files in use properly
 *    2. calls into
 *        LogMessage()
 *        ObjFactory()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. all temp, log and XML files are closed out properly
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns status from either WshShell.run or ReadFileContents
 *********************************************************************/
function CloseFiles()
{
    EnterFunction( arguments );

    var iRtn;

    //don"t want to miss any data
    if( 0 < g_oObjects.oIsaXml.childNodes.length )
    {
        SaveXML( g_oObjects.oIsaXml );
        CleanupXML( );
    }

    LogMessage( g_oMessages.L_AllDone_txt + new Date().toString() );
    if( g_oVariables.fDebugMode )
    {
        LogMessage( GetData() );
    }
    if ( g_oObjects.fsTraceFile != null )
    {
        try
        {
            g_oObjects.fsTraceFile.Close();
            g_oObjects.fsTraceFile = null;
        }
        catch( err )
        {
        //since file logging fails, we need to leave a record in the event log
            g_oObjects.oWsh.LogEvent( 1, g_oMessages.L_SaveFailed_txt + 
                    g_oVariables.szXmlFile + 
                    g_oMessages.L_ErrNum_txt + ToHex( err.number ) + 
                    g_oMessages.L_ErrDesc_txt + err.description +
                    g_oMessages.L_CopyMsg_txt );
            ExitFunction( arguments, ToHex( err.number ) );
            return err;
        }
    }

    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;

}

function GetData()
{
    var szRtn = "\r\n++++++++++++++++++++++\r\nContents of the variables data set:\r\n++++++++++++++++++++++\r\n";
    for( var szItem in g_oVariables )
    {
        szRtn += Fprintf( "g_oVariables.%1 == %2\r\n", new Array( szItem, g_oVariables[ szItem ] ) );
    }
    return szRtn;
}

/**********************************************************************
 * SaveXML( XmlDomDoc )
 * This function:
 *    1. saves the "raw" XML
 *    2. calls into
 *        LogMessage()
 *  3. called by 
 *        - several functions -
 *
 * if successful:
 *    Current XML data is saved to g_oVariables.szXmlFile
 *    returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    errors are reported to the log and optionally to the user
 *    returns error number
 *********************************************************************/
function SaveXML( XmlDomDoc )
{
    EnterFunction( arguments );

    LogMessage( g_oMessages.L_SaveFile_txt + g_oVariables.szXmlFile );
    try
    {
        XmlDomDoc.save( g_oVariables.szXmlFile );
    }
    catch( err )
    {
        ShowXmlError( XmlDomDoc, g_oMessages.L_SaveFile_txt + 
                    g_oVariables.szXmlFile );
        ExitFunction( arguments, ToHex( err.number ) );
        return err.number;
    }
    ExitFunction( arguments, g_oVariables.lS_OK );
    return g_oVariables.lS_OK;
}

/**********************************************************************
 * CleanupXML( )
 * This function:
 *    1. transforns the "raw" XML from isainfo additions to nicely 
 *        indented XML like that produced by ISA
 *    2. calls into
 *        ShowXmlError()
 *  3. called by 
 *        CloseFiles()
 *
 * if successful:
 *    All XML is indented nicely instead of being a single string
 *
 * if unsuccessful:
 *    errors are reported to the log and optionally to the user
 *********************************************************************/
function CleanupXML( )
{
    EnterFunction( arguments );

    var oWriter = ObjFactory( "Msxml2.MXXMLWriter.3.0" );
    oWriter.indent = true;                    //makes the output more readable
    oWriter.encoding = "UTF-8";
    oWriter.version = "1.0";
    //temporary XmlDomDocument
    var oXmlDomDoc = ObjFactory( "Msxml2.DomDocument.3.0" );
    oWriter.output = oXmlDomDoc;
    var oReader = ObjFactory( "Msxml2.SAXXMLReader.3.0" );
    oReader.contentHandler = oWriter;
    oReader.errorHandler = oWriter;
    oReader.dtdHandler = oWriter;
    oReader.putProperty( "http://xml.org/sax/properties/lexical-handler", oWriter );
    oReader.putProperty( "http://xml.org/sax/properties/declaration-handler", oWriter );

    LogMessage( g_oMessages.L_Cleanup_txt );
    /*
     * parse the existing XML structure with oWriter (handler for oReader)
     */
    try
    {
        oReader.parse( g_oObjects.oIsaXml );
    }
    catch( err )
    {
        ShowXmlError( oWriter, g_oMessages.L_ParseFailed_txt );
        ExitFunction( arguments, ToHex( err.number ) );
        return err.number;
    }

    oWriter.output = "" // required to actually acquire the reformatted output in some cases.
    if( 0 == oXmlDomDoc.xml.length )
    {
        LogMessage( "g_oObjects.oIsaXml parsing produced 0-length output from oWriter" );
        return  g_oVariables.lS_OK;
    }
    return SaveXML( oXmlDomDoc );
}


/**********************************************************************
 * ReadFileContents( oParentNode, szFilePath )
 * This function:
 *    1. read from the file at "szFilePath"
 *        append the data into a text node as an oParentNode child element
 *        delete the temp file
 *    2. calls into
 *        LogMessage()
 *        LogError()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. all temp, log and XML files are closed out properly
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns status from either WshShell.run or ReadFileContents
 *********************************************************************/
function ReadFileContents( oParentNode, szFilePath, bLogErrorText )
{
    EnterFunction( arguments );

    var fsTempFile = null;
    var iForReading = 1;
    
    LogMessage( g_oMessages.L_GetFile_txt + szFilePath );
    try
    { 
        fsTempFile = g_oObjects.oFSO.OpenTextFile( szFilePath, iForReading );
    }
    catch( err )
    {
        if( bLogErrorText )
        {
            var oCdataNode = oParentNode.ownerDocument.createCDATASection( g_oMessages.L_GetFileFailed_txt + szFilePath );
            oParentNode.appendChild( oCdataNode );
            return true;
        }
        LogError( err, g_oMessages.L_GetFileFailed_txt + szFilePath );
        ExitFunction( arguments, ToHex( err.number ) );
        return false;
    }

    var szTempData = fsTempFile.readAll();
    var oCdataNode = oParentNode.ownerDocument.createCDATASection( szTempData );
    oParentNode.appendChild( oCdataNode );
    fsTempFile.Close();
    return true;

}

/**********************************************************************
 * GetIsaFilesData( oParentNode, szFolderPath )
 * This function:
 *    1. opens the the folder at "szFolderPath" using g_oObjects.oFSO
 *        enumerates the subfolder objects within
 *    2. calls into
 *        GetExeFilesInfo()
 *        GetIsaFilesData()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. returns oParentNode populated with data from GetExeFilesData()
 *
 * if unsuccessful:
 *    1. oFolderNode contents depend on failure location
 *********************************************************************/
function GetIsaFilesData( oParentNode, szRootPath, szServerName )
{
    EnterFunction( arguments );

    var szFolderPath;
    var oThisFolder;
    var oFolderNode;
    var szSubFolders = new Array( "\\clients", 
                        "\\clients\\Program Files\\Microsoft Firewall Client 2004", 
                        "\\MSDE", "\\Uninstall" );
    var szMsfwRegVal        = "HKLM\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\" +
                                "MSFW\\Setup\\SQLPath";
    var szMsfwLocation = "";        //ISA MSDE installation path
    var szNotFound = "80070002";

    oFolderNode = NewChildNode( oParentNode, "Folder", "" );
    NewChildNode( oFolderNode, "Path", szRootPath );
    GetExeFilesInfo( oFolderNode, szRootPath, false );

    for( var iInx in szSubFolders )
    {
        szFolderPath = szRootPath + szSubFolders[ iInx ];
        oFolderNode = NewChildNode( oParentNode, "Folder", "" );
        NewChildNode( oFolderNode, "Path", szFolderPath );
        GetExeFilesInfo( oFolderNode, szFolderPath, false );
    }

    try
    {
        szMsfwLocation = g_oObjects.oWsh.RegRead( szMsfwRegVal ) + "\\Binn";
        oFolderNode = NewChildNode( oParentNode, "Folder", "" );
        NewChildNode( oFolderNode, "Path", szMsfwLocation );
        GetExeFilesInfo( oFolderNode, szMsfwLocation, false );
    }
    catch( err )
    {
        if( ToHex( err.number ) != szNotFound )
        {
            LogError( err, "trying to get the MSFW registry settings.")
        }
        else
        {
            err.clear;
        }
    }
}

/**********************************************************************
 * GetExeFilesInfo( oParentNode, oFsoFolder )
 * This function:
 *    1. Enumerates the executable files in oFsoFolder
 *        gathers data related to those files
 *    2. calls into
 *        GetWmiObjectsInfo()
 *        g_oObjects.oFSO.GetFileVersion()
 *  3. called by 
 *        GetIsaFilesData()
 *
 * if successful:
 *    1. returns oParentNode populated with file-specific data
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetExeFilesInfo( oParentNode, szFolderPath )
{
    EnterFunction( arguments );

    var szQuery = "select AccessMask, Compressed, Encrypted, " +
                        "FileSize, FileType, " +
                        "Manufacturer, Name from CIM_DataFile where ";
    var szFileQuery2 = "Drive = \"";
    var szFileQuery3 = "\" and Path = \"";
    var szFileQuery4 = "\" and (Extension = \"exe\" or Extension " +
                        "= \"dll\" or Extension = \"sys\")";
    var cFiles;
    var szDrive;
    var szPath;
    var szQuery;
    var szFileVer;
    
    
    szDrive = szFolderPath.substr( 0, szFolderPath.indexOf( "\\" ) );
    szPath = szFolderPath.substr( szFolderPath.indexOf( "\\" ) ) + "\\";
    szPath = szPath.replace( /\\/g, "\\\\" );
    szQuery += ( szFileQuery2 + szDrive + szFileQuery3 + szPath + szFileQuery4 );

    GetWmiObjectsInfo( oParentNode, szQuery, "", "File" );
}

/*#######################################
 # end of file functions
 ######################################*/


/*#######################################
 # Windows Management Interface support 
 # functions
 ######################################*/
    /*#######################################
     # WMI registry support functions
     ######################################*/
/**********************************************************************
 * EnumRegistryKeys( oParentNode, lHive, szKeyPath )
 * This function:
 *    1. enumerates registry keys starting from szKeyPath
 *    2. calls into
 *        EnumRegistryValues()
 *        EnumRegistryKeys()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. returns oParentNode populated with registry data
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function EnumRegistryKeys( oParentNode, lHive, szKeyPath )
{
    EnterFunction( arguments );

    var arrKeys;
    var szSubKey;
    var szSubKeys;
    var oMethod;
    var oOutParam;
    var oInParam;
    var iInx;

    var oKeyNode = NewChildNode( oParentNode, "RegKey", "" );
    var oNewNode = NewChildNode( oKeyNode, "Name", szKeyPath );

    LogMessage( g_oMessages.L_EnumKey_txt + szKeyPath );
    EnumRegistryValues( oKeyNode, lHive, szKeyPath );

    oMethod = g_oObjects.oWmiReg.Methods_.Item( "EnumKey" );
    oInParam = oMethod.InParameters.SpawnInstance_(); 
    oInParam.hDefKey = lHive;
    oInParam.sSubKeyName = szKeyPath;

    if( 0 < szKeyPath.indexOf( "fpc\\storage" ) )
    {
        return true;   //will use "tier" filtering later
    }
    try
    { 
        oOutParam = g_oObjects.oWmiReg.ExecMethod_( oMethod.Name, oInParam );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_EnumKeysFail_txt + szKeyPath );
        ExitFunction( arguments, false );
        return false;
    }
    if( null != oOutParam.sNames )
    {
        arrKeys = oOutParam.sNames.toArray();
        LogMessage("arrKeys is " + arrKeys.length + " items long.");
        for ( var iInx in arrKeys )
        {
            EnumRegistryKeys( oKeyNode, lHive, szKeyPath + "\\" + arrKeys[ iInx ] );
        }
    }
    return true;
}

/**********************************************************************
 * EnumRegistryValues( oParentNode, lHive, szKeyPath )
 * This function:
 *    1. enumerates registry keys starting from szKeyPath
 *        populates oParentNode with registry value data
 *    2. calls into
 *        GetRegistryValue()
 *  3. called by 
 *        EnumRegistryKeys()
 *
 * if successful:
 *    1. returns oParentNode populated with registry data
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function EnumRegistryValues( oParentNode, lHive, szKeyPath, oFilter )
{
    EnterFunction( arguments );

    var oNewNode;
    var oMethod;
    var oOutParam;
    var oInParam;
    var arrValues;
    var vValue;
    var iInx;
    var arrTypes;
    var arrTypeNames = new Array( "", "String", "ExpandedString", "Binary", 
                            "Dword", "", "", "MultiString" );

    ShowStatus( "." );
    var oMethod = g_oObjects.oWmiReg.Methods_.Item( "EnumValues" );
    var oInParam = oMethod.InParameters.SpawnInstance_(); 
    oInParam.hDefKey = lHive;
    oInParam.sSubKeyName = szKeyPath;

    try
    {
        oOutParam = g_oObjects.oWmiReg.ExecMethod_( oMethod.Name, oInParam );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_EnumValuesFail_txt + szKeyPath );
        ExitFunction( arguments, false );
        return false;
    }        

    if( null == oOutParam.sNames )
    {
        ExitFunction( arguments, true );
        return true;
    }
    arrValues = oOutParam.sNames.toArray();
    arrTypes = oOutParam.Types.toArray();
    for ( iInx in arrValues )
    {
        var oValueNode = NewChildNode( oParentNode, "RegVal", "" );
        
        NewChildNode( oValueNode, "Name", arrValues[ iInx ] );
        NewChildNode( oValueNode, "Type", arrTypeNames[ arrTypes[ iInx ] ] );

        vValue = GetRegistryValue( arrTypeNames[ arrTypes[ iInx ] ], lHive, 
                                szKeyPath, arrValues[ iInx ] );
        NewChildNode( oValueNode, "Value", vValue );
        ShowStatus(".");
    }
    ExitFunction( arguments, true );
    return true;
}

/**********************************************************************
 * GetRegistryValue( szValueType, lHive, szKey, szValueName )
 * This function:
 *    1. populates oParentNode with registry value data
 *    2. calls into
 *        g_oObjects.oWmiReg.ExecMethod_()
 *        ToHex()
 *  3. called by 
 *        EnumRegistryValues()
 *
 * if successful:
 *    1. returns oParentNode populated with registry data
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetRegistryValue( szValueType, lHive, szKey, szValueName )
{
    EnterFunction( arguments );

    var oMethod;
    var oInParam; 
    var oOutParam;
    var vRegValue = null;
    var iInx;
    
    oMethod = g_oObjects.oWmiReg.Methods_.Item( "Get" + szValueType + "Value" );
    oInParam = oMethod.InParameters.SpawnInstance_(); 
    oInParam.hDefKey = lHive;
    oInParam.sSubkeyName = szKey;
    oInParam.sValueName = szValueName;

    try
    { 
        oOutParam = g_oObjects.oWmiReg.ExecMethod_( oMethod.Name, oInParam );
    }
    catch( err )
    {
        LogError( err, szKey + "\\" + ToHex( oOutParam ) );
        vRegValue = "Not Found";
        ExitFunction( arguments, vRegValue );
        return vRegValue;
    }

    switch( szValueType )
    {
        case "String":
        case "ExpandedString":
            if ( oOutParam.sValue != null )
            {
                vRegValue = oOutParam.sValue;
            }
            break;
         case "MultiString":
            if ( oOutParam.sValue != null )
            {
                vRegValue = oOutParam.sValue.toArray();
            }
            break;
        case "Binary":
            if ( oOutParam.uValue != null )
            {
                vRegValue = oOutParam.uValue.toArray();
            }
            break;
        case "Dword":
            if ( oOutParam.uValue != null )
            {
                vRegValue = ToHex( oOutParam.uValue );
            }
    }
    
    return vRegValue;

}

    /*#######################################
     # end of WMI registry functions
     ######################################*/

    /*#######################################
     # WMI generic support functions
     ######################################*/
/**********************************************************************
 * GetWmiObjects( szServerName )
 * This function:
 *    1. attempts to connect to the root\cimv2 and root\default
 *        namespaces on the specified server
 *    2. calls into
 *        GetObject()
 *        LogError()
 *        LogMessage()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. g_oObjects.oWmiCimv2 and g_oObjects.oWmiReg are set to the respective objects on 
 *        the server specified by szServerName
 *    2. returns "true"
 *
 * if unsuccessful:
 *    1. g_oObjects.oWmiCimv2 and g_oObjects.oWmiReg are set to null
 *    2. returns "false"
 *********************************************************************/
function GetWmiObjects( szServerName )
{
    EnterFunction( arguments );

    LogMessage( g_oMessages.L_AccessWMI_txt + szServerName );

    //try to access the WMI cimv2 namespace on the specified server
    try
    {
        g_oObjects.oWmiCimv2 = GetObject( "winmgmts:{ImpersonationLevel=Impersonate, " +
                                          "(security, backup)}!\\\\" + szServerName + 
                                          "\\root\\cimv2" );
    }
    catch( err )
    {
        g_oObjects.oWmiCimv2 = null;
        LogError( err, g_oMessages.L_NoWmiCimv2_txt + szServerName  );
        ExitFunction( arguments, false );
        return false;
    }

    //try to access the local WMI StdRegProv class in the default namespace on the specified server
    try
    {
        g_oObjects.oWmiReg = GetObject( "winmgmts:{ImpersonationLevel=Impersonate, " +
                                        "(security, backup)}!\\\\" + szServerName + 
                                        "\\root\\default:StdRegProv" );
    }
    catch( err )
    {
        g_oObjects.oWmiReg = null;
        LogError( err, g_oMessages.L_NoWmiReg_txt + szServerName );
        ExitFunction( arguments, false );
        return false;
    }
    return true;
}

/**********************************************************************
 * GetWmiObjectsInfo( oParentNode, szQuery, szSet, szPart )
 * This function:
 *    1. abstracts the gathering of SWbemObject data
 *    2. calls into
 *        LogMessage()
 *        g_oObjects.oWmiCimv2.ExecQuery()
 *        LogError()
 *        GetWmiObjectProperties()
 *  3. called by 
 *        GetOSData()
 *        GetHardwareData()
 *        GetEvtLogData()
 *        GetExeFilesInfo()
 *
 * if successful:
 *    1. oParentNode is populated with data according to the query passed
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *    2. returns the err object
 *********************************************************************/
function GetWmiObjectsInfo( oParentNode, szQuery, szSet, szPart )
{
    EnterFunction( arguments );

    LogMessage( g_oMessages.L_WmiQuery_txt + szQuery );
    var wbemFlagReturnImmediately = 0x10;
    var wbemFlagForwardOnly = 0x20;
    var sWbemFlags = wbemFlagReturnImmediately + wbemFlagForwardOnly;
    var oGroupNode = oParentNode;
    if ( szSet != "" )
    {
        oGroupNode = NewChildNode( oParentNode, szSet, "" );
    }
        
    var SWbemObjectSet = null;
    try
    {
        SWbemObjectSet = g_oObjects.oWmiCimv2.ExecQuery( szQuery, "WQL", sWbemFlags );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_QueryFailed_txt + szQuery );
        ExitFunction( arguments, ToHex( err.number ));
        return err;
    }
    
    var cItems = new Enumerator( SWbemObjectSet );
    var inx = 0;
    for ( ; !cItems.atEnd(); cItems.moveNext() )
    {
        inx++;
        var oItem = cItems.item();
        var oItemNode = NewChildNode( oGroupNode, szPart, "" );
        GetWmiObjectProperties( oItemNode, szPart, oItem )
        ShowStatus( "." );
        if( "LogEntry" == szPart && 100 == inx)
        {
            break;
        }
    }

    LogMessage( "\r\n -- " +  inx + g_oMessages.L_QueryResults_txt );

    SaveXML( g_oObjects.oIsaXml );
    return g_oVariables.lS_OK;
}


/**********************************************************************
 * GetWmiObjectProperties( oParentNode, szQuery, szSet, szPart )
 * This function:
 *    1. abstracts the gathering of SWbemObject property data
 *    2. calls into
 *        LogMessage()
 *        LogError()
 *        GetDiskPartition()
 *        GetExtraDrvSvcData()
 *        GetProcessUserData()
 *        GetShareSecurity()
 *  3. called by 
 *        GetWmiObjectsInfo()
 *
 * if successful:
 *    1. oParentNode is populated with properties data
 *    2. returns nothing
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *    2. returns the err object
 *********************************************************************/
function GetWmiObjectProperties( oParentNode, szPart, SWbemObject )
{
    EnterFunction( arguments );

    var cProperties = new Enumerator( SWbemObject.properties_ );
    for ( ; !cProperties.atEnd(); cProperties.moveNext() )
    {
        var oProperty = cProperties.item();
        var vPropValue = "";
        if( null != oProperty.value )
        {
            vPropValue = oProperty.value;
            if( oProperty.isArray )
            {
                vPropValue = oProperty.value.toArray();
            }
        }

        NewChildNode( oParentNode, oProperty.name, vPropValue );
    }
    /*
     * Some items need additional info
     */
    switch( szPart )
    {
        case "Drive":
            GetDiskPartition( oParentNode, SWbemObject );
            break;
        case "Driver":
        case "Service":
            GetExtraDrvSvcData( oParentNode, SWbemObject.Name );
            break;
        case "Process":
            GetProcessUserData( oParentNode, SWbemObject );
            break;
        case "Share":
            GetShareSecurity( oParentNode, SWbemObject );
            break;
        case "File":
            var szFileVer = g_oObjects.oFSO.GetFileVersion( SWbemObject.Name );
            NewChildNode( oParentNode, "Version", szFileVer );
            
    }

}

/**********************************************************************
 * GetDiskPartition( oParentNode, Disk )
 * This function:
 *    1. associates a logical disk with its physical drive and partition
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        GetWmiObjectsInfo()
 *
 * if successful:
 *    1. creates a new oParentNode childnode as "Partition", containing 
 *        data from the physical location of the logical disk
 *    2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. returns err from call to GetWmiObjectsInfo()
 *    2. oParentNode child node contents are dependent on the failure point
 *********************************************************************/
function GetDiskPartition( oParentNode, oDisk )
{
    EnterFunction( arguments );

    var oNewNode;
    var oPartNode;
    var cPhysDisks;
    var oPhysDisk;
    var szProps = new Array( "BootPartition", "Description", 
                            "DeviceID", "PrimaryPartition" );
    var iInx;
    var vTempData;
    
    cPhysDisks = new Enumerator( oDisk.Associators_( "", 
                                "Win32_DiskPartition" ) );

    for( ; !cPhysDisks.atEnd(); cPhysDisks.moveNext() )
    {
        oPhysDisk = cPhysDisks.item();
        oNewNode = oParentNode.ownerDocument.createElement( "Partition" );
        oPartNode = oParentNode.insertBefore( oNewNode, 
                    oParentNode.selectSingleNode( "Size" ) );
        
        for( iInx in szProps )
        {
            vTempData = eval( "oPhysDisk." + szProps[ iInx ] );
             oNewNode = NewChildNode( oPartNode, szProps[ iInx ], vTempData );
        }
    }
}

/**********************************************************************
 * GetExtraDrvSvcData( oParentNode, szName )
 * This function:
 *    1. Gathers service and driver dependency information
 *    2. calls into
 *        GetRegistryValue()
 *  3. called by 
 *        GetWmiObjectsInfo()
 *
 * if successful:
 *    1. oParentNode is populated with the dependecies as listed in the 
 *        registry for the item specified in szName
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *********************************************************************/
function GetExtraDrvSvcData( oParentNode, szName )
{
    EnterFunction( arguments );

    var szRtn;
    var oNewNode;
    var szDependsKey = "SYSTEM\\CurrentControlSet\\Services\\";
    var szDepends = new Array( "DependOnService", "DependOnGroup" );
    var iInx;
    var oDescrNode;

    oDescrNode = oParentNode.selectSingleNode( "Description" );

    for( iInx in szDepends )
    {
        oNewNode = oParentNode.ownerDocument.createElement( szDepends[ iInx ] );
        oParentNode.insertBefore( oNewNode, oDescrNode );
        szRtn = GetRegistryValue( "MultiString", g_oVariables.lHKLM, szDependsKey + szName, 
                                szDepends[ iInx ] );
        if ( szRtn != null )
        {
            oNewNode.text = szRtn;
        }
    }
}

/**********************************************************************
 * IsLocalRras( )
 * This function:
 *    1. Determins if RRAS is installed and running on the specified 
 *        server
 *    2. calls into
 *        LogMessage()
 *        g_oObjects.oWmiCimv2.ExecQuery()
 *        LogError()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. returns "true" if RRAS is installed and running,"false" otherwise
 *
 * if unsuccessful:
 *    1. returns "false"
 *********************************************************************/
function IsLocalRras( )
{
    EnterFunction( arguments );

    var oRRASSvc;
    var szRRASQuery = "select Name, Started, StartMode from Win32_Service " +
                    "where Name=\"RemoteAccess\"";
    var SWbemObjectSet;
    
    LogMessage( g_oMessages.L_CheckRras_txt );
    try
    {
        SWbemObjectSet = new Enumerator( g_oObjects.oWmiCimv2.ExecQuery( szRRASQuery ) );
        for ( ; !SWbemObjectSet.atEnd(); SWbemObjectSet.moveNext() )
        {
            oRRASSvc = SWbemObjectSet.item();
            if ( ( oRRASSvc.StartMode == "Auto" || oRRASSvc.StartMode == "Manual" ) &&
                oRRASSvc.Started == true )
            {
                return true;
            }
        }
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_NoRras_txt );
        ExitFunction( arguments, ToHex( err.number ) );
    }
    return false;

}

/**********************************************************************
 * GetWmiDate( iOffset )
 * This function:
 *    1. Calculates the WMI-fomat for the date plus any offset (+ or -)
 *    2. calls into
 *        - none -
 *  3. called by 
 *        GetEvtLogData()
 *
 * if successful:
 *    1. returns an offset date string in WMI format:
 *    yyyymmddhhmmss.000000sOOO
 *
 * if unsuccessful:
 *    1. no special logic
 *********************************************************************/
function GetWmiDate( iOffset )
{
    EnterFunction( arguments );

    var Now            = null;
    var diffDate    = dateVal;
    var dateVal;
    var WmiDate;
    var WmiMonth;
    var WmiDay;
    var WmiHour;
    var WmiMin;
    var WmiSec;
    var WmiDate;
    var Offset;
    var dayMilli    = 86400000;    //number of mS in a day (1000 * 3600 * 24)

    Now                = new Date();
    dateVal            = Date.parse( Now.toString() );
    
    if( iOffset != 0 )
    {
        diffDate = (iOffset < 0 )?
                    new Date( dateVal - ( dayMilli * Math.abs(iOffset) ) ):
                    new Date( dateVal + ( dayMilli * Math.abs(iOffset) ) );
    }

    dateVal = diffDate.getMonth() + 1;
    WmiMonth = ( dateVal < 10 )? "0" + dateVal: dateVal;
    dateVal = diffDate.getDate();
    WmiDay = ( dateVal < 10 )? "0" + dateVal: dateVal;
    dateVal = diffDate.getHours();
    WmiHour = ( dateVal < 10 )? "0" + dateVal: dateVal;
    dateVal = diffDate.getMinutes(); 
    WmiMin = ( dateVal < 10 )? "0" + dateVal: dateVal;
    dateVal = diffDate.getSeconds(); 
    WmiSec = ( dateVal < 10 )? "0" + dateVal: dateVal;
    Offset = diffDate.getTimezoneOffset();
    Offset = ( Offset < 0 )? "+" + Math.abs( Offset ).toString(): 
                "-" + Math.abs( Offset ).toString();

    WmiDate = diffDate.getFullYear().toString() +
            WmiMonth +
            WmiDay +
            WmiHour +
            WmiMin +
            WmiSec +
            ".000000" +
            Offset;
    
    return WmiDate;

}
    /*#######################################
     # End of WMI generic support functions
     ######################################*/

    /*#######################################
     # WMI Share enumeration functions
     ######################################*/
/**********************************************************************
 * GetShares( oParentNode )
 * This function:
 *    1. Enumerates the shares on a server using the Win32_Share class
 *    2. calls into
 *        GetWmiObjectsInfo()
 *  3. called by 
 *        GetPermissions()
 *
 * if successful:
 *    1. oParentNode is populated with data representing the shares on
 *        the server and the relevant permissions
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *********************************************************************/
function GetShares( oParentNode )
{
    EnterFunction( arguments );

    var szShareQuery = "Select AllowMaximum, Caption, Description, " +
                        "MaximumAllowed, Name, Path, Status, Type " +
                        "From Win32_Share";

    try
    {
        GetWmiObjectsInfo( oParentNode, szShareQuery, "Win32_Share", "Share" );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_QueryFailed_txt + " Win32_Share." );
        ExitFunction( arguments, ToHex( err.number ));
    }
}

/**********************************************************************
 * GetShareSecurity( oShareNode, oShare )
 * This function:
 *    1. Gathers the security settings for the share and related folder
 *    2. calls into
 *        g_oObjects.oWmiCimv2.Get()
 *        GetWmiAcls()
 *  3. called by 
 *        GetPermissions()
 *
 * if successful:
 *    1. oParentNode is populated with data representing the security
 *        settings for the share and related folder if appropriate
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *********************************************************************/
function GetShareSecurity( oShareNode, oShare )
{
    EnterFunction( arguments );

    var oNewNode = null;
    var oFolderNode = null;
    var szSharePath = oShare.Path.toString();
    var szShareName = oShare.Name.toString();
    var oShareAcls = null;
    var oMethod = null;
    var oOutParam = null;
    var lAdmin = 0x80000000;             //denotes an admin share
    var lWmiNotFound = -2147217406;     //0x80041002 - object not found error
    var lDrive = 0;                        //drive share type

    if ( ( oShare.Type & lAdmin ) != lAdmin )
    {
        try
        {
            oShareAcls = g_oObjects.oWmiCimv2.Get( "Win32_LogicalShareSecuritySetting.Name=\"" + 
                        szShareName + "\"" );
            oMethod = oShareAcls.Methods_( "GetSecurityDescriptor" );
            oOutParam = oShareAcls.ExecMethod_( oMethod.Name );
            GetWmiAcls( oShareNode, oOutParam.Descriptor );
        }
        catch( err )
        {
            /*
             * if "not found" returned, it"s null SecDescr (Everyone=full)
             */
            if ( err.number != lWmiNotFound )
            {
                LogError( err, g_oMessages.L_QueryFailed_txt + szShareName );
                ExitFunction( arguments, ToHex( err.number ));
                return;
            }
            err.clear;
        }

    }

    if ( ( oShare.Type & 0xF ) == lDrive )
    {
        oFolderNode = NewChildNode( oShareNode, "Folder", "" );
        NewChildNode( oFolderNode, "Path", szSharePath );
        GetFileSecurity( oFolderNode, szSharePath );
    }

}

/**********************************************************************
 * GetFileSecurity( oParentNode, szFilePath )
 * This function:
 *    1. Gathers the security settings for the share and related folder
 *    2. calls into
 *        g_oObjects.oWmiCimv2.Get()
 *        GetWmiAcls()
 *  3. called by 
 *        GetShareSecurity()
 *        GetWmiObjectsInfo()
 *
 * if successful:
 *    1. oParentNode is populated with data representing the security
 *        settings for the share and related folder if appropriate
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *********************************************************************/
function GetFileSecurity( oParentNode, szFolderPath )
{
    EnterFunction( arguments );

    var oFolder = null;
    var oOutParam = null;
    var lWmiNotFound = -2147217406;     //0x80041002; object not found error

    try
    {
        oFolder = g_oObjects.oWmiCimv2.Get( "Win32_LogicalFileSecuritySetting.Path=\"" + 
                    szFolderPath.replace( /\\/, "\\\\" ) + "\"" );
        oOutParam = oFolder.ExecMethod_( "GetSecurityDescriptor" );
        GetWmiAcls( oParentNode, oOutParam.Descriptor );
    }
    catch( err )
    {
        /*
         * if "not found" returned, it"s either a "no disc in drive" response
         * or the shared resource was removed (folder deleted, etc.).  Either
         * way, it"s not critical.
         */
        if ( err.number == lWmiNotFound )
        {
            NewChildNode( oParentNode, "SecurityDescriptor", err.description )
        }
        else
        {
            LogError( err, g_oMessages.L_AccessFailed_txt + szFolderPath );
        }
        ExitFunction( arguments, ToHex( err.number ));
    }
}


/**********************************************************************
 * GetWmiAcls( oParentNode, oWmiSecDescr )
 * This function:
 *    1. Enumerates DACL and SACL collections contained in oWmiSecDescr
 *    2. calls into
 *        GetWmiAces()
 *  3. called by 
 *        GetShareSecurity()
 *
 * if successful:
 *    1. oParentNode is populated with data representing the DACL and SACL
 *        settings found in oWmiSecDescr
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on the failure location
 *********************************************************************/
function GetWmiAcls( oParentNode, oWmiSecDescr )
{
    EnterFunction( arguments );

    var oNewNode = null;
    var oAclNode = null;
    var oOwner = null;
    var oGroup = null;
    var oAcl = null;
    var szTemp = "";

    oAclNode = NewChildNode( oParentNode, "SecurityDescriptor", "" );

    NewChildNode( oAclNode, "ControlFlags", oWmiSecDescr.ControlFlags );
    
    oOwner = oWmiSecDescr.Owner;
    GetCredentials( oAclNode, oOwner, "Owner" );

    oGroup = oWmiSecDescr.Group;
    GetCredentials( oAclNode, oGroup, "Group" );

    oNewNode = NewChildNode( oAclNode, "DACL", "" );
    oAcl = oWmiSecDescr.DACL.toArray();
    GetWmiAces( oNewNode, oAcl );

    oNewNode = NewChildNode( oAclNode, "SACL", "" );
    if ( oWmiSecDescr.SACL != null )
    {
        oAcl = oWmiSecDescr.SACL.toArray();
        GetWmiAces( oNewNode, oAcl );
    }
}

/**********************************************************************
* GetWmiAces( oParentNode,oWmiAcl )
* This function:
*    1. Enumerates ACE collections contained in oWmiAcl and gathers
*        the data they contain
*    2. calls into
*        -nothing-
*  3. called by 
*        GetWmiAcls()
*
* if successful:
*    1. oParentNode is populated with data representing the ACE data 
*        contained in oWmiAcl
*
* if unsuccessful:
*    1. oParentNode contents depend on the failure location
*********************************************************************/
function GetWmiAces( oParentNode, oWmiAcl )
{
    EnterFunction( arguments );

    var oNewNode;
    var oAceNode;
    var oTrusteeNode;
    
    var oWmiAce;
    var iInx;
    var oTrustee;
    
    for ( iInx in oWmiAcl )
    {
        oWmiAce = oWmiAcl[ iInx ];
        oAceNode = NewChildNode( oParentNode, "ACE", "" );

        NewChildNode( oAceNode, "Mask", oWmiAce.AccessMask );
        NewChildNode( oAceNode, "Type", oWmiAce.AceType );

        /*
         * some predefined groups have no "domain"
         */
        oTrustee = oWmiAce.Trustee;
        if( oTrustee != null )
        {
            oTrusteeNode = NewChildNode( oAceNode, "Trustee", "" );

            oNewNode = NewChildNode( oTrusteeNode, "Domain", "" );
            if ( oTrustee.Domain != null )
            {
                oNewNode.text = oTrustee.Domain;
            }

            oNewNode = NewChildNode( oTrusteeNode, "Name", "" );
            if ( oTrustee.Name != null )
            {
                oNewNode.text = oTrustee.Name;
            }

            oNewNode = NewChildNode( oTrusteeNode, "SID", "" );
            if ( oTrustee.SIDString != null )
            {
                oNewNode.text = oTrustee.SIDString;
            }
        }
    }
}

/**********************************************************************
 * GetCredentials( oParentNode, oWmiSecDescr )
 * This function:
 *    1. Creates new XMLDomElements as children of oParentNode and sets
 *        the text element to appropriate values
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        GetWmiAcls()
 *
 * if successful:
 *    1. adds new children of oParentNode
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function GetCredentials( oParentNode, oAclCreds, szContext )
{
    EnterFunction( arguments );

    var szTemp = "";

    var oNewNode = NewChildNode( oParentNode, szContext, "" );
    if( oAclCreds != null )
    {
        if ( oAclCreds.Domain != null )
        {
            szTemp = oAclCreds.Domain;
        }
        NewChildNode( oNewNode, "Domain", szTemp );
        szTemp = "";

        if ( oAclCreds.Name != null )
        {
            szTemp = oAclCreds.Name;
        }
        NewChildNode( oNewNode, "Name", szTemp );
    }

}

    /*#######################################
     # End of WMI Share enumeration functions
     ######################################*/

    /*#######################################
     # IIS Support functions
     ######################################*/

/**********************************************************************
 * CheckIIS( oParentNode, szServerName )
 * This function:
 *    1. Reports the active IIS services
 *    2. calls into
 *        EnumIISServers()
 *        NewChildNode()
 *  3. called by 
 *        CheckIIS()
 *
 * if successful:
 *    1. returns at least two new child nodes of oParentNode as:
 *        "Name" - Virtual Server "friendly name"
 *        "State" - Current operational state of this server
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function CheckIIS( oParentNode, szServerName )
{
    EnterFunction( arguments );

    LogMessage( g_oMessages.L_GetIIS_txt );

    var arrServices = new Array( "W3SVC", "MSFTPSVC", "SMTPSVC", "NNTPSVC" );
    var oIisNode = NewChildNode( oParentNode, "IIS", "" );
    var iNotInstalled = -2147024893;
    var iSvcCount = 0;
    szServerName += "/";

    try
    {
        GetObject( "IIS://" );
    }
    catch( err )
    {
        err.clear;
        oIisNode.text = "not installed";
        LogMessage( iSvcCount + g_oMessages.L_FoundIIS_txt );
        return;
    }

    for( var inx in arrServices )
    {
        var oSvcsNode = NewChildNode( oIisNode, arrServices[ inx ], "" );
        var szConnection = "IIS://" + szServerName + arrServices[ inx ];
        var cServices = null;
        try
        {
            cServices = GetObject( szConnection );
        }
        catch( err )
        {
            if( err.number == iNotInstalled )
            {
                err.clear;
                oSvcsNode.text = "not installed";
                continue;
            }
            var szErr = "Failed to connect to " + szConnection;
            LogError( err, szErr );
            ExitFunction( arguments, ToHex( err.number ));
        }
        iSvcCount++
        NewChildNode( oSvcsNode, "DisableSocketPooling", cServices.DisableSocketPooling );
        EnumIISServers( oSvcsNode, cServices );
    }

    LogMessage( iSvcCount + g_oMessages.L_FoundIIS_txt );

}

/**********************************************************************
 * EnumIISServers( oParentNode, cServers )
 * This function:
 *    1. Reports the server name and current state
 *    2. calls into
 *        NewChildNode()
 *        GetIisBindings()
 *  3. called by 
 *        CheckIIS()
 *
 * if successful:
 *    1. returns at least two new child nodes of oParentNode as:
 *        "Name" - Virtual Server "friendly name"
 *        "State" - Current operational state of this server
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function EnumIISServers( oParentNode, cServers )
{
    EnterFunction( arguments );

    var eServers = new Enumerator( cServers );
    //
    //  Values for MD_SERVER_STATE from iiscnfg.h
    //
    var MD_SERVER_STATES = new Array( "", "Starting", "Started", "Stopping",
                                    "Stopped", "Pausing", "Paused", "Continuing"
                                    );

    for( ; !eServers.atEnd(); eServers.moveNext() )
    {
        var oServer = eServers.item();
        //only numeric vservers have bindings
        if( !isNaN( oServer.Name ) )
        {
            var oServerNode = NewChildNode( oParentNode, "VirtualServer", "");
            NewChildNode( oServerNode, "Name", oServer.ServerComment );
            NewChildNode( oServerNode, "State", MD_SERVER_STATES[ oServer.ServerState ] );
            NewChildNode( oServerNode, "DisableSocketPooling", oServer.DisableSocketPooling );
            GetIisBindings( oServerNode, oServer.ServerBindings.toArray() );
            GetIisBindings( oServerNode, oServer.SecureBindings.toArray() );
        }
    }
}

/**********************************************************************
 * GetIisBindings( oParentNode, arrBindings )
 * This function:
 *    1. Reports the bindings associated with the array of serverbindings
 *    2. calls into
 *        NewChildNode()
 *  3. called by 
 *        EnumIISServers()
 *
 * if successful:
 *    1. returns three new child nodes of oParentNode as:
 *        "IP" - listening IP address
 *        "Name" - host header name
 *        "Port" - listening port
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function GetIisBindings( oParentNode, arrBindings )
{
    EnterFunction( arguments );

    var inx = 0;
    for( inx in arrBindings )
    {
        var oBindNode = NewChildNode( oParentNode, "Binding", "" );
        var arrBinding = arrBindings[ inx ].split( ":" );
        var szIP = arrBinding[ 0 ];
        if( szIP.length == 0 )
        {
            szIP = "All unassigned";
        }

        NewChildNode( oBindNode, "IP", szIP );
        NewChildNode( oBindNode, "Name", arrBinding[ 2 ] );
        NewChildNode( oBindNode, "Port", arrBinding[ 1 ] );
    }
}


    /*#######################################
     # End of IIS Support functions
     ######################################*/

/**********************************************************************
 * GetProcessUserData( oParentNode, oProcess )
 * This function:
 *    1. Obtains the user acct that a specific process is running in
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        GetWmiObjectsInfo
 *
 * if successful:
 *    1. returns two child nodes of oParentNode as:
 *        "Owner" - domain\username format
 *        "OwnerSID - Win32_SIDString format
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function GetProcessUserData( oParentNode, oProcess )
{
    EnterFunction( arguments );

    var oOutParam;
    var szUser;

    try
    {
        oOutParam = oProcess.ExecMethod_( "GetOwner" );
    }
    catch( err )
    {
        NewChildNode( oParentNode, "Owner", "Error getting owner data." );
        NewChildNode( oParentNode, "OwnerSID", "" );
        ExitFunction( arguments, ToHex( err.number ));
        err.clear;
        return;
    }

    if( oOutParam.Domain != null  )
    {
        szUser = oOutParam.Domain + "\\" + oOutParam.User;
    }
    else if( oOutParam.User != null )
    {
        szUser = oOutParam.User;
    }
    else
    {
        szUser = "";
    }
    NewChildNode( oParentNode, "Owner", szUser );
    
    oOutParam = oProcess.ExecMethod_( "GetOwnerSID" );

    if( oOutParam.Sid != null  )
    {
        szUser = oOutParam.Sid;
    }
    else
    {
        szUser = "";
    }
    NewChildNode( oParentNode, "OwnerSID", szUser );
        
}
/*#######################################
 # End of Windows Management Interface 
 # support functions
 ######################################*/

/*#######################################
 # Basic helper functions
 ######################################*/
/**********************************************************************
 * NewChildNode( oParentNode, szNodeName, szNodeText )
 * This function:
 *    1. Creates a new XMLDomElement as a child of oParentNode and sets
 *        the text element to sznodeText
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        - just about everyone -
 *
 * if successful:
 *    1. returns oNewNode as a child of oParentNode
 *
 * if unsuccessful:
 *    - none -
 *********************************************************************/
function NewChildNode( oParentNode, szNodeName, vNodeText )
{
    EnterFunction( arguments );

    var oNewNode  = null;
    var szNodeText = vNodeText.toString();
    try
    {
        oNewNode = oParentNode.ownerDocument.createElement( szNodeName );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_CreateFailed1_txt + szNodeName + 
                        g_oMessages.L_CreateFailed2_txt + szNodeText +
                        g_oMessages.L_CreateFailed3_txt );
        ExitFunction( arguments, null );
    }

    oParentNode.appendChild( oNewNode );
    /*
     * have to use CDATA to avoid XML parsing errors in the HTA
     */
    if( szNodeText )
    {
        var oCdataNode = oParentNode.ownerDocument.createCDATASection( szNodeText );
        oNewNode.appendChild( oCdataNode );
    }
    return oNewNode;
}


/*#######################################
 # End of Basic helper functions
 ######################################*/

/*#######################################
 # ISA-specific helper functions
 ######################################*/
/**********************************************************************
 * ExportIsaXml( )
 * This function:
 *    1. attempts to exercise the Export() method to gather ISA
 *        configuration
 *    2. calls into
 *        Array.Export() ISA COM method to obtain the ISA backup data
 *  3. called by Main()
 *
 * if successful:
 *    1. returns g_oVariables.lS_OK
 *    2. ISA XML configuration is read into g_oObjects.oIsaXml
 *
 * if unsuccessful:
 *    1. returns error generated by the call to Array.Export()
 *    2. g_oObjects.oIsaXml contents are dependent on the failure point
 *********************************************************************/
function ExportIsaXml( )
{
    EnterFunction( arguments );

    var oContext = g_oObjects.oThisArray;   //SE export location
    var iOptionalData = 6;               //everything except passwords
    var szComment         = g_oMessages.L_Comment1_txt + 
                            g_oVariables.szThisUser + 
                            g_oMessages.L_Comment2_txt +
                            " as \"" + g_oMessages.szScriptName +
                            g_oMessages.szCmdOpts + "\"";


    LogMessage( g_oMessages.L_ReadIsaData_txt );

    if( g_oVariables.fServerOnly )
    {
        GetFakeXml( null, szComment );
        return GetIsa2K4ServerData( 
                g_oObjects.oIsaXml.selectSingleNode( "//fpc4:Server" ) );
    }

    if( g_oVariables.fEntMode )
    {
        oContext = g_oObjects.oISA;
    }

    try
    {
        oContext.Export( g_oObjects.oIsaXml, 
                        iOptionalData, 
                        "", 
                        szComment );
        SaveXML( g_oObjects.oIsaXml );
    }
    catch( err )
    {
        GetFakeXml( err, szComment );
        LogError( err, g_oMessages.L_IsaDataFailed_txt );
        return GetIsa2K4ServerData( 
                g_oObjects.oIsaXml.selectSingleNode( "//fpc4:Server" ) );
    }
    
    g_oVariables.fSaveXml = true;
    return GetIsa2K4ArraysData( g_oObjects.oIsaXml.selectSingleNode( "//fpc4:Arrays" ) );
}

/**********************************************************************
 * GetIsa2K4ArraysData( oParentNode, oArrays )
 * This function:
 *    1. Optionally Enumerates the arrays collection or a specific array
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4SignaledAlerts()
 *  3. called by 
 *        GetIsa2K4ArraysData()
 *
 * if successful:
 *    1. oParentNode contains data from selected ISA Arrays Properties
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4ArraysData( oParentNode )
{
    EnterFunction( arguments );

    var oArrays = null;
    var oArray = null;
    var oArrayNode = null;

    if( !g_oVariables.fEE )
    {
        g_oObjects.oIsaArray = g_oObjects.oThisArray;
        oArrayNode = oParentNode.selectSingleNode( "fpc4:Array" );
        return GetIsa2K4ArrayData( oArrayNode );
    }

    if( g_oVariables.fOneArray )
    {
        if( GetIsaArrayObject( g_oVariables.szIsaArray ) == false )
        {
            return !g_oVariables.lS_OK;
        }

        oArrayNode = 
            oParentNode.selectSingleNode( "fpc4:Array[@StorageName = \"" + 
                            g_oObjects.oIsaArray.PersistentName + "\"]" );
        return GetIsa2K4ArrayData( oArrayNode );
    }

    oArrays = new Enumerator( g_oObjects.oISA.Arrays );
    for( ; !oArrays.atEnd(); oArrays.moveNext() )
    {
        g_oObjects.oIsaArray = oArrays.item();
        oArrayNode = 
            oParentNode.selectSingleNode( "fpc4:Array[@StorageName = \"" +
                            g_oObjects.oIsaArray.PersistentName + "\"]" );
        if( GetIsa2K4ArrayData( oArrayNode ) != g_oVariables.lS_OK )
        {
            return !g_oVariables.lS_OK;
        }
    }
    return g_oVariables.lS_OK;
}

/**********************************************************************
 * GetIsa2K4ArrayData( oParentNode )
 * This function:
 *    1. Gathers additional Array data
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4SignaledAlerts()
 *  3. called by 
 *        GetIsa2K4ArrayData()
 *
 * if successful:
 *    1. oParentNode contains data from selected ISA Array Properties
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4ArrayData( oParentNode )
{
    EnterFunction( arguments );

    var oServersNode = null;
    var oLogsNode = null;

    oLogsNode = oParentNode.selectSingleNode( "fpc4:Logs" );
    if( GetIsa2K4LoggingData( oLogsNode, g_oObjects.oIsaArray.Logging )
        != g_oVariables.lS_OK )
    {
        return !g_oVariables.lS_OK;
    }

    SaveXML( g_oObjects.oIsaXml );
    oServersNode = oParentNode.selectSingleNode( "fpc4:Servers" );
    return GetIsa2K4ServersData( oServersNode );

}

/**********************************************************************
 * GetIsa2K4ServersData( oParentNode )
 * This function:
 *    1. Gathers additional Array data
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4SignaledAlerts()
 *  3. called by 
 *        GetIsa2K4ArrayData()
 *
 * if successful:
 *    1. oParentNode contains data from selected ISA Array Properties
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4ServersData( oParentNode )
{
    EnterFunction( arguments );

    var oServers = null;
    var oServer = null;
    var oServerNode = null;

    if( !g_oVariables.fEE )
    {
        g_oObjects.oIsaServer = g_oObjects.oThisServer;
        oServerNode = oParentNode.selectSingleNode( "fpc4:Server" );
        return GetIsa2K4ServerData( oServerNode );
    }

    if( g_oVariables.fOneServer )
    {
        if( GetIsaServerObject( g_oVariables.szIsaServer ) == false )
        {
            return !g_oVariables.lS_OK;
        }

        oServerNode = 
            oParentNode.selectSingleNode( "fpc4:Server[@StorageName = \"" + 
                            g_oObjects.oIsaServer.PersistentName + "\"]" );
        return GetIsa2K4ServerData( oServerNode );
    }

    oServers = new Enumerator( g_oObjects.oIsaArray.Servers );
    for( ; !oServers.atEnd(); oServers.moveNext() )
    {
        g_oObjects.oIsaServer = oServers.item();
        oServerNode = 
            oParentNode.selectSingleNode( "fpc4:Server[@StorageName = \"" + 
                            g_oObjects.oIsaServer.PersistentName + "\"]" );

        if( GetIsa2K4ServerData( oServerNode ) != g_oVariables.lS_OK )
        {
            return !g_oVariables.lS_OK;
        }
    }
}

/**********************************************************************
 * GetIsa2K4ServerData( oParentNode, oServer )
 * This function:
 *    1. Gathers additional ISA-specific Server data
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4SignaledAlerts()
 *  3. called by 
 *        GetIsa2K4ServersData()
 *
 * if successful:
 *    1. oParentNode contains data from selected ISA Server Properties
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4ServerData( oParentNode )
{
    EnterFunction( arguments );

    if( g_oObjects.oIsaServer == null )
    {
        if( g_oVariables.fAdminOnly )
        {
            g_oObjects.oIsaServer = g_oObjects.oThisArray.Servers.Item( 1 );
        }
        else
        {
            g_oObjects.oIsaServer = g_oObjects.oISA.GetContainingServer();
        }
    }
    
    var iInx;
    var oServerNode;
    var oNewNode;
    var szDataSet    = new Array( "CreatedTime", "Description", 
                    "FirewallServiceStatus", "FirewallServiceUpTime", 
                    "FQDN", "InstallationDirectory", 
                    "JobSchedulerServiceStatus", "JobSchedulerServiceUpTime", 
                    "MSDEServiceStatus", "ProductID", "ProductVersion", 
                    "ServerStatus" );
    var NotInstalled = "80070424";
    var oServer = g_oObjects.oIsaServer;

    LogMessage( g_oMessages.L_IsaSvrData_txt );

    oServerNode = NewChildNode( oParentNode, "ISAInfoData", "" );
    NewChildNode( oServerNode, "ScriptVersion", g_oMessages.L_Ver_txt );

    for ( iInx in szDataSet )
    {
        ShowStatus( "." );
        oNewNode = NewChildNode( oServerNode, szDataSet[ iInx ], "" );
        try
        {
            oNewNode.text = eval( "oServer." + szDataSet[ iInx ] );
        }
        catch( err )
        {
            switch( ToHex( err.number ) )
            {
                case NotInstalled:
                    if( szDataSet[ iInx ] == "MSDEServiceStatus" )
                    {
                        oNewNode.text = NotInstalled;
                        err.clear;
                    }
                    break;
                default:
                    LogError( err, g_oMessages.L_AccessFailed_txt + 
                                    "oServer." + szDataSet[ iInx ] + 
                                    " on " + g_oObjects.oIsaServer.Name );
            }
        }
    }

    LogMessage( "" )
    SaveXML( g_oObjects.oIsaXml );
    GetIsa2K4SignaledAlerts( oServerNode );
    return GetServerData( oServerNode );
}

/**********************************************************************
 * GetIsa2K4SignaledAlerts( oParentNode, oServer )
 * This function:
 *    1. Enumerates ISA Server Signaled alerts
 *    2. calls into
 *        LogMessage()
 *        GetAlertInstances()
 *  3. called by 
 *        GetIsa2K4ServerData()
 *
 * if successful:
 *    1. oParentNode contains data from ISA Server Signaled Alerts
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4SignaledAlerts( oParentNode )
{
    EnterFunction( arguments );

    LogMessage( g_oMessages.L_SigAlerts_txt );

    var oNewNode;
    var oAlertsNode;
    var oAlertNode;
    var cSignaledAlerts;
    var oSignaledAlert;
    var oAlertItem;
    var i;
    
    var szAlertData = new Array( "AdditionalKey", "Count", "Name", "Server", 
                                "Severity" );

    cSignaledAlerts = g_oObjects.oIsaServer.SignaledAlerts;

    oAlertsNode = NewChildNode( oParentNode, "SignaledAlerts", "" );    
    for ( i = 1; i <= cSignaledAlerts.Count; i++ )
    {
        ShowStatus( "." );
        oSignaledAlert = cSignaledAlerts( i );
        oAlertNode = NewChildNode( oAlertsNode, "Alert", "" );
        
        for ( oAlertItem in szAlertData )
        {
            oNewNode = NewChildNode( oAlertNode, szAlertData[ oAlertItem ], "" );
            try
            {
                oNewNode.text = eval( "oSignaledAlert." + szAlertData[ oAlertItem ] );
            }
            catch( err )
            {
                ExitFunction( arguments, ToHex( err.number ) );
                LogError( err, g_oMessages.L_AccessFailed_txt + " SignaledAlert." + 
                            szAlertData[ oAlertItem ] + " on " + 
                            g_oObjects.oIsaServer.Name );
            }
        }

        GetAlertInstances( oAlertNode, g_oObjects.oIsaServer.Name, oSignaledAlert );
    }

    WScript.StdOut.WriteLine( );
    SaveXML( g_oObjects.oIsaXml );
    return g_oVariables.lS_OK;
}

/**********************************************************************
 * GetAlertInstances( oParentNode, oServer )
 * This function:
 *    1. Enumerates ISA Server Signaled Alert instances
 *    2. calls into
 *        LogMessage()
 *  3. called by 
 *        GetIsa2K4SignaledAlerts()
 *
 * if successful:
 *    1. oParentNode contains data from ISA Server Signaled Alert Instances
 *     2. returns g_oVariables.lS_OK
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetAlertInstances( oParentNode, szServerName, oSignaledAlert )
{
    EnterFunction( arguments );

    var oNewNode;
    var oInstanceNode;
    var oSignaledAlert;
    var oAlertItem;
    var cInstances;
    var oInstance;
    var oInstanceItem;
    var i;
    
    var szAlertInstance = new Array( "Acknowledged", "Count", "Description", 
                                "Resolution", "TimeStamp" );

    for ( i = 1; i <= oSignaledAlert.Count; i++ )
    {
        oInstance = oSignaledAlert( i );
        oInstanceNode = NewChildNode( oParentNode, "Instance", "" );
        
        for ( oInstanceItem in szAlertInstance )
        {
            oNewNode = NewChildNode( oInstanceNode, 
                                    szAlertInstance[ oInstanceItem ], "" );
            try
            {
                oNewNode.text = eval( "oInstance." + 
                                        szAlertInstance[ oInstanceItem ] );
            }
            catch( err )
            {
                ExitFunction( arguments, ToHex( err.number ) );
                LogError( err, g_oMessages.L_AccessFailed_txt + " AlertInstance." + 
                            szAlertInstance[ oInstanceItem ] + 
                            " in " + oSignaledAlert.Name + 
                            " on " + szServerName );
            }
        }
    }

    return g_oVariables.lS_OK;
}

/**********************************************************************
 * GetIsa2K4LoggingData( oParentNode, oLogs )
 * This function:
 *    1. Gathers additional ISA-specific Array logging settings
 *    2. calls into
 *        LogMessage()
 *        GetIsa2K4SignaledAlerts()
 *  3. called by 
 *        GetServerData()
 *
 * if successful:
 *    1. oParentNode contains data from selected ISA Server Properties
 *     2. returns "g_oVariables.lS_OK"
 *
 * if unsuccessful:
 *    1. oParentNode contents depend on failure location
 *********************************************************************/
function GetIsa2K4LoggingData( oParentNode, oLogs )
{
    EnterFunction( arguments );

    var iInx;
    var cLogs;
    var oLog;
    var oLogsNode;
    var oLogNode;
    var oNewNode;
    var szDataSet = new Array( "DeleteOldLogsOnLimitExceeded", 
        "KeepMinFreeDiskSpace", "LimitTotalLogSize", "LogDbTableName", 
        "LogDbUserName", "LogFileKeepOld", "MaxTotalLogSizeGB", 
        "MinFreeDiskSpaceMB" );
    var vTempData;

    LogMessage( g_oMessages.L_LogConfig_txt );

    NewChildNode( oParentNode, "MSDENumberOfInsertsPerBatch", 
                                oLogs.MSDENumberOfInsertsPerBatch );
    NewChildNode( oParentNode, "MSDEQueryTimeout", 
                                oLogs.MSDEQueryTimeout );

    cLogs = new Enumerator( oLogs );
    for( ; !cLogs.atEnd(); cLogs.moveNext() )
    {
        oLog = cLogs.item();
        oLogNode = oParentNode.selectSingleNode( 
                "fpc4:Log[@StorageName=\"" + oLog.PersistentName + "\"]" );
        for( iInx in szDataSet )
        {
            try
            {
                vTempData = eval( "oLog." + szDataSet[ iInx ] );
                NewChildNode( oLogNode, szDataSet[ iInx ], vTempData );
            }
            catch( err )
            {
                LogError( err, g_oMessages.L_AccessFailed_txt + " oLog." + szDataSet[ iInx ] );
            }
        }
    }

    return g_oVariables.lS_OK;
}

/**********************************************************************
 * GetIsaArrayObject( szArrayName )
 * This function:
 *    1. wraps the "Getserver" method with a try-catch and input validation
 *    2. calls into
 *        LogMessage()
 *        LogError()
 *  3. called by 
 *        GetIsa2xArrayData()
 *
 * if successful:
 *    1. g_oObject.oIsaArray = the selected array
 *     2. returns true
 *
 * if unsuccessful:
 *    1. reports error and returns false
 *********************************************************************/
function GetIsaArrayObject( szArrayName )
{
    EnterFunction( arguments );

    if( szArrayName == "" )
    {
        szArrayName = g_oObjects.oThisArray.Name;
    }

    try
    {
        g_oObjects.oIsaArray = g_oObjects.oISA.Arrays( szArrayName );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_AccessFailed_txt + "Array \"" + 
            szArrayName + "\"");
        ExitFunction( arguments, false );
        return false;
    }
    
    return true;
}

/**********************************************************************
 * GetIsaServerObject( szServerName )
 * This function:
 *    1. wraps the "GetServer" method with a try-catch and input validation
 *    2. calls into
 *        LogMessage()
 *        LogError()
 *  3. called by 
 *        GetIsa2xServerData()
 *
 * if successful:
 *    1. g_oObject.oIsaServer = the selected Server
 *     2. returns true
 *
 * if unsuccessful:
 *    1. reports error and returns false
 *********************************************************************/
function GetIsaServerObject( szServerName )
{
    EnterFunction( arguments );

    if( szServerName == "" )
    {
        szServerName = g_oObjects.oThisServer.Name;
    }

    try
    {
        g_oObjects.oIsaServer = g_oObjects.oIsaArray.Servers( szServerName );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_AccessFailed_txt + "Server \"" + 
            szServerName + "\"");
        ExitFunction( arguments, false );
        return false;
    }
    
    return true;
}


/**********************************************************************
 * GetISA( )
 * This function:
 *    1. Creates the default ISA COM object
 *    2. calls into
 *        LogError()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. g_oObjects.oISA is set to a valid ISA COM object
 *     2. returns true
 *
 * if unsuccessful:
 *    1. LogError indicate the failure and cause
 *    2. returns false
 *********************************************************************/
function GetISA( )
{
    EnterFunction( arguments );

     try
    {
        var szSE = "FPC.Root";
        g_oObjects.oISA = ObjFactory( szSE );
        return true;
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_GetISAFailed_txt )
        ExitFunction( arguments, false );
        return false;
    }
}

/**********************************************************************
 * SortItOut( )
 * This function:
 *	1. determines if the local host is:
 *		ISA 2000 or ISA 2004
 *		SE or EE
 *		Admin or Server
 *	2. calls into
 *      CheckIsa2k()
 *      CheckIsa2k4()
 *  3. called by 
 *		Main()
 *
 * if successful:
 *	1. ISA environment is spelled out in three variables:
 *		g_oVariables.fIsa2k
 *		g_oVariables.fEE
 *		g_oVariables.fAdminOnly
 *
 * if unsuccessful:
 *	1. errors are handled in called functions
 *********************************************************************/
function SortItOut( )
{
	CheckIsaVersion();
    CheckIsaEdition();
    CheckIsaAdminOnly();

    if( g_oVariables.fIsa2k )
    {
        LogError( null, "ISAInfo only runs on an ISA Server version 2004 or greater." );
        return false;
    }

    var fRtn = false;
    if( g_oVariables.fServerOnly )
    {
        fRtn = true;
    }
    else if( g_oVariables.fEE )
    {
        fRtn = ConnectToCss( g_oVariables.szCssName, g_oVariables.szCssUser, g_oVariables.szCssDomain, g_oVariables.szCssPass );
    }
    else 
    {
        fRtn = ConnectToArray( g_oVariables.szArrayName, g_oVariables.szCssUser, g_oVariables.szCssDomain, g_oVariables.szCssPass );
    }
    ExitFunction( arguments, fRtn );
    return fRtn;
}

/**********************************************************************
 * CheckIsa2k()
 * This function:
 *	1. determines if the local host is:
 *		ISA 2000
 *		SE or EE
 *	2. calls into
 *		LogError()
 *		GetIsa2KEE()
 *		Check2kAdmin()
 *  3. called by 
 *		DetermineEnvironment()
 *
 * if successful:
 *	1. Sets the state of three variables:
 *		g_oVariables.fIsa2k
 *		g_oVariables.fEE
 *		g_oVariables.fAdminOnly
 *
 * if unsuccessful:
 *	1. shows any unexpected error
 *********************************************************************/
function CheckIsaVersion()
{
 	try
	{
		g_oObjects.oWsh.RegRead( g_oVariables.szAcmeComp );
		g_oVariables.fIsa2k = true;
	}
	catch( err )
	{
		/*
		 * if trying to read g_oVariables.szAcmeComp == err, 
		 * then this isn"t ISA2K
		 */
		err.clear;
		g_oVariables.fIsa2k = false;
	}
}

function CheckIsaEdition()
{
 	try
	{
		var szEE = g_oObjects.oWsh.RegRead( 
						g_oVariables.szIsaEdition );
		g_oVariables.fEE = 
				( szEE == g_oVariables.szEEGuid )? true: false;
	}
	catch( err )
	{
        err.clear;
        g_oVariables.fEE = false;
	}
}

function CheckIsaAdminOnly()
{
    try
    {
    	var szTest = g_oObjects.oWsh.RegRead( g_oVariables.szArrayGuid );
        g_oVariables.fAdminOnly = false;
    }
    catch( err )
    {
        err.clear;
        g_oVariables.fAdminOnly = true;
    }
}
/**********************************************************************
 * ConnectToCss()
 * This function:
 *    1. attempts to connect to the default CofigurationStorageServer
 *        Admin or Server
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckIsa2k()
 *
 * if successful:
 *    1. Sets the state of g_oVariables.fAdminOnly
 *
 * if unsuccessful:
 *    1. Sets the state of g_oVariables.fAdminOnly
 *********************************************************************/
function ConnectToCss( szCss, szUsername, szDomain, szPassword )
{
    EnterFunction( arguments );

    if( "" == szCss )
    {
        try
        {
            szCss = g_oObjects.oISA.ConfigurationStorageServer;
        }
        catch( err )
        {
            LogError( err, "Failed to determine Configuration Storage Server name and none supplied." );
            ExitFunction( arguments, false );
            return false;
        }
    }
    LogMessage( Fprintf( "Connecting to CSS \"%1\"", new Array( szCss ) ) );
    try
    {
        g_oObjects.oISA.ConnectToConfigurationStorageServer( szCss, szUsername, szDomain, szPassword );
    }
    catch( err )
    {
        LogError( err, Fprintf( "Failed to connect to CSS \"%1\" using Username (\"%2\") Domain (\"%3\") and Password (\"%4\").",
                            new Array( szCss, szUsername, szDomain, szPassword ) ) );
        ExitFunction( arguments, false );
        return false;
    }

    //default to current array if none specified
    if( g_oVariables.fAdminOnly && "" == g_oVariables.szIsaArray )
    {
        LogMessage( "User failed to supply an array name or index.." );
        ExitFunction( arguments, false );
        return false;
    }
    if( !g_oVariables.fAdminOnly &&  "" == g_oVariables.szIsaArray )
    {
        LogMessage( "Connecting to \"Containing\" Array.." );
        g_oObjects.oThisArray = g_oObjects.oIsaArray = g_oObjects.oISA.GetContainingArray();
        g_oVariables.szIsaArray = g_oObjects.oThisArray.Name;
        ExitFunction( arguments, true );
        return true;
    }
    //can't connect to a string representation of a number
    if( !isNaN( g_oVariables.szIsaArray ) )
    {
        g_oVariables.szIsaArray = parseInt( g_oVariables.szIsaArray );
    }
    LogMessage( Fprintf( "Connecting to Array \"%1\"", new Array( g_oVariables.szIsaArray.toString() ) ) );
    try
    {
        g_oObjects.oThisArray = g_oObjects.oIsaArray = g_oObjects.oISA.Arrays.Item( g_oVariables.szIsaArray );
        g_oVariables.szIsaArray = g_oObjects.oThisArray.Name;
        ExitFunction( arguments, true );
        return true;
    }
    catch( err )
    {
        LogError( err, Fprintf( "Failed to connect to ISA Array \"%1\".", new Array( g_oVariables.szIsaArray ) ) );
        ExitFunction( arguments, false );
        return false;
    }
}

/**********************************************************************
 * ConnectToArray()
 * This function:
 *    1. attempts to connect to the specified Array
 *        Admin or Server
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckIsa2k()
 *
 * if successful:
 *    1. Sets the state of g_oVariables.fAdminOnly
 *
 * if unsuccessful:
 *    1. Sets the state of g_oVariables.fAdminOnly
 *********************************************************************/
function ConnectToArray( szUsername, szDomain, szPassword )
{
    EnterFunction( arguments );

    if( "" == g_oVariables.szIsaArray )
    {
        if( g_oVariables.fAdminOnly )
        {
            LogError( null, "** Failed to specify the Array name." );
            return false;
        }
        g_oObjects.oIsaArray = g_oObjects.oISA.GetContainingArray();
        g_oObjects.oThisArray = g_oObjects.oIsaArray;
        return true;
    }
    LogMessage( Fprintf( "Connecting to Array \"%1\"", new Array( g_oVariables.szIsaArray ) ) );
    try
    {
        g_oObjects.oIsaArray = g_oObjects.oISA.Arrays.Connect( g_oVariables.szIsaArray, szUsername, szDomain, szPassword );
        g_oObjects.oThisArray = g_oObjects.oIsaArray;
        return true;
    }
    catch( err )
    {
        LogError( err, Fprintf( "Failed to connect to Array \"%1\" using Username (\"%2\") Domain (\"%3\") and Password (\"%4\").",
                            new Array( g_oVariables.szIsaArray, szUsername, szDomain, szPassword ) ) );
        return false;
    }
}

/**********************************************************************
 * DetermineIsaEnvironment( )
 * This function:
 *    1. determines if the local host is:
 *        ISA 2000 or ISA 2004
 *    2. calls into
 *        GetISA()
 *      CheckIsa2k()
 *      CheckIsa2k4()
 *        LogError()
 *  3. called by 
 *        Main()
 *
 * if successful:
 *    1. ISA environment is spelled out in three variables:
 *        g_oVariables.fIsa2k
 *        g_oVariables.fSE
 *        g_oVariables.fAdminOnly
 *
 * if unsuccessful:
 *    1. errors are handled in called functions
 *********************************************************************/
function DetermineIsaEnvironment( )
{
    EnterFunction( arguments );

    if ( false == GetISA() )
    {
        return false;
    }

    return SortItOut();
}

/*#######################################
 # End of ISA-specific helper functions
 ######################################*/

/**********************************************************************
 * GetFakeXml()
 * This function:
 *    1. Attempts to create the "root" XML needed if the ISA.export() 
 *        method fails or if /serveronly is used
 *    2. calls into
 *        - none -
 *  3. called by 
 *        ExportIsaXml()
 *
 * if successful:
 *    1. creates some XML to 
 *
 * if unsuccessful:
 *    1. returns null
 *    2. displays the error
 *    3. throws the error back to the caller
 *********************************************************************/
function GetFakeXml( oErr, szComment )
{
    EnterFunction( arguments );

    var szFakeXML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<fpc4:Root StorageName=\"FPC\" " +
                   "xmlns:fpc4=\"http://schemas.microsoft.com/isa/config-4\">" +
                   "<fpc4:Comment>" + szComment;
    if( oErr != null )
    {
        szFakeXML += "&lt;br&gt;" +
                   "&lt;span style=&quot;color:red;font-weight:bold&quot;&gt;" +
                   "ISA Backup failed&lt;br&gt;" + 
                   "Error = 0x" + ToHex( oErr.number ) + "; " + oErr.description +
                   "&lt;/span&gt;";
    }
    szFakeXML +=   "</fpc4:Comment>" +
                   "<fpc4:Arrays><fpc4:Array><fpc4:Servers><fpc4:Server>" +
                   "<fpc4:Name>" + g_oVariables.szThisServer + "</fpc4:Name>" +
                   "</fpc4:Server></fpc4:Servers></fpc4:Array></fpc4:Arrays>" +
                   "<fpc4:Enterprise/></fpc4:Root>";

    if( !g_oObjects.oIsaXml.loadXML( szFakeXML ) )
    {
        ShowXmlError( g_oObjects.oIsaXml, "failed to load " + szFakeXML);
        WScript.Quit();
    }
    g_oVariables.fSaveXml = true;
}

/**********************************************************************
 * ObjFactory( szNewObjName )
 * This function:
 *    1. Attempts to create the ActiveX Object specified by szNewObjName
 *        supports remote when szNewObjName is constructed as:
 *        "ObjName, ServerName" if the COM object supports remoting AND
 *        is registered on the local machine
 *    2. calls into
 *        LogError()
 *  3. called by 
 *        Main()
 *        InitEnvironment()
 *        DoSystemCmd()
 *        CloseFiles()
 *
 * if successful:
 *    1. returns the desired object
 *
 * if unsuccessful:
 *    1. returns null
 *    2. displays the error
 *    3. throws the error back to the caller
 *********************************************************************/
function ObjFactory( szNewObjName )
{
    EnterFunction( arguments );

    var oNewObject = null;

    if ( g_oObjects.fsTraceFile != null )
    {
        LogMessage( g_oMessages.L_AccessObj_txt + szNewObjName + "\"" );
    }

    try
    {
        oNewObject = new ActiveXObject( szNewObjName );
    }
    catch( err )
    {
        LogError( err, g_oMessages.L_ObjCreateErr_txt + szNewObjName );
        throw err;
    }
    return oNewObject;
}


/**********************************************************************
 * ToHex( lValue)
 * This function:
 *    1. Converts a number to its hexadecimal equivalent and accounts for 
 *        negative numbers (hResults)
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        - nearly all functions -
 *
 * errors are not evaluated
 *********************************************************************/
function ToHex( lValue )
{
    EnterFunction( arguments );

    lValue = ( lValue < 0 )? lValue + 0x100000000: lValue;
    return lValue.toString( 16 ).toUpperCase();
}

/**********************************************************************
 * Fprintf( szMessage, arrVariables )
 * This function:
 *    1. Replaces data from arrVariables into szMessage
 *        
 *    2. calls into
 *        - none -
 *
 *  3. called by 
 *        - nearly all functions -
 *
 *    errors are not evaluated
 *********************************************************************/
function Fprintf( szMessage, arrVariables )
{
    for( var inx in arrVariables )
    {
        var oRegEx = new RegExp( "\%" + ( parseInt( inx ) + 1 ), "g" )
        szMessage = szMessage.replace( oRegEx,
                                       arrVariables[ inx ]
                                      );
    }
    return szMessage;
}

/**********************************************************************
 * ShowStatus( szMessage )
 * This function:
 *    1. Sends szMessage to WScript.StdOut
 *    2. calls into
 *        - nothing -
 *  3. called by 
 *        - nearly all functions -
 *
 * errors are not evaluated
 *********************************************************************/
function ShowStatus( szMessage )
{
    EnterFunction( arguments );

    WScript.StdOut.Write( szMessage );
}

/**********************************************************************
 * LogMessage( szMessage )
 * This function:
 *    1. Writes szMessage to the trace log and the console if not running
 *        under MPSReports
 *    2. calls into
 *        g_oObjects.fsTraceFile.Write()
 *  3. called by 
 *        - nearly all functions -
 *
 * if successful:
 *    1. trace log and optionally the console contain szMessage
 *
 * if unsuccessful:
 *    1. depends on the error
 *********************************************************************/
function LogMessage( szMessage )
{
    /*
     * don"t need to spew debugging info to the screen...
     */
    if( !g_oVariables.fMPSReports && 
        szMessage.indexOf( "-->" ) == -1 )
    {
        WScript.Echo( szMessage );
    }

    /*
     * don"t timestamp entries beginning with "\r\n"
     */
    if( szMessage.substr( 0, 2 ) != "\r\n" )
    {
        szMessage = "[ " + new Date().toTimeString() + " ] " + szMessage;
    }

    if ( g_oObjects.fsTraceFile )
    {
        g_oObjects.fsTraceFile.WriteLine( szMessage );
    }

}

/**********************************************************************
 * LogError( oErr, szMessage )
 * This function:
 *    1. Displays szMessage and any error data if not running in MPSReports
 *        
 *    2. calls into
 *        LogMessage
 *  3. called by 
 *        - nearly all functions -
 *
 *    errors are not evaluated
 *********************************************************************/
function LogError( oErr, szMessage )
{
    EnterFunction( arguments );
    
    var Exclamation = 48;
    var YesNo = 4;
    var Yes = 6;
    var No = 7;
    var RtnVal;

    if( null != oErr )
    {
        szMessage += g_oMessages.L_ErrNum_txt + ToHex( oErr.number ) + 
                g_oMessages.L_ErrDesc_txt + oErr.description;
        oErr.clear;
    }

    LogMessage( szMessage );

    if ( g_oVariables.fQuiet )
    {
        return;
    }
    szMessage += g_oMessages.L_Continue_txt ;
    RtnVal = g_oObjects.oWsh.Popup( szMessage + g_oMessages.L_CopyMsg_txt, 
                            0, g_oMessages.L_TitleMsg_txt, 
                            Exclamation + YesNo );
    if( RtnVal == No )
    {
        LogMessage( g_oMessages.L_Canceled_txt );
        CloseFiles();
        WScript.quit();
    }
}


/**********************************************************************
 * ShowXmlError( oXml, szMessage )
 * This function:
 *    1. Displays szMessage and any error data if not running in MPSReports
 *        
 *    2. calls into
 *        LogMessage
 *  3. called by 
 *        - nearly all functions -
 *
 *    errors are not evaluated
 *********************************************************************/
function ShowXmlError( oXml, szMessage )
{
    EnterFunction( arguments );

    var WshShell = ObjFactory( "WScript.Shell" );
    var oXmlError = oXml.parseError;
    var Exclamation = 48;
    var YesNo = 4;
    var Yes = 6;
    var No = 7;
    var RtnVal;
    
    if( null != oXmlError )
    {
        szMessage += (g_oMessages.L_ErrNum_txt + ToHex( oXmlError.errorCode ) + 
                      g_oMessages.L_ErrDesc_txt + oXmlError.reason + 
                      g_oMessages.L_Err_Line_txt + oXmlError.line + 
                      g_oMessages.L_Err_Char_txt + oXmlError.linepos +
                      g_oMessages.L_Err_Text_txt + oXmlError.srcText
                      );
    }

    LogError( null, szMessage );
}


/**********************************************************************
 * ParseArgs( )
 * This function:
 *    1. Controls evaluation of the cmd-line arguments
 *    2. calls into
 *      CheckArgument()
 *  3. called by 
 *        main()
 *
 * if successful:
 *    1. user-specified options are set
 *  2. returns true
 *
 * if unsuccessful:
 *    1. returns false
 *********************************************************************/
function ParseArgs( )
{
    EnterFunction( arguments );

    var oArgs = WScript.Arguments;
    var inx;
    
    /*
     * valid arguments are listed in ISAInfoVariables and ISAInfoMessages
     */
    for( inx = 0; inx < oArgs.length; inx++ )
    {
        g_oMessages.szCmdOpts += ( " " + oArgs( inx ) );
    }

    if( oArgs.length > g_oVariables.iMaxArgs )
    {
        return false;
    }
    
    for( inx = 0; inx < oArgs.length; inx++ )
    {
        if( CheckArgument( oArgs( inx ) ) == false )
        {
            return false;
        }
    }
    return true;
}


/**********************************************************************
 * CheckArgument( )
 * This function:
 *    1. Tests each cmd-line argument and sets appropriate values
 *        tests for duplicate options
 *    2. calls into
 *        - none -
 *  3. called by 
 *        main()
 *
 * if successful:
 *    1. specified property is set
 *     2. returns true
 *
 * if unsuccessful or duplicate option:
 *    1. returns false
 *********************************************************************/
function CheckArgument( szArgument )
{
    EnterFunction( arguments );

    var iEndVal1 = szArgument.indexOf( ":" );
    var szArg = "";
    var szValue = "";
    
    if( iEndVal1 == -1 ) //":" not found
    {
         szArg = szArgument;
    }
    else
    {
         szArg = szArgument.substr( 0, iEndVal1 );
         szValue = szArgument.substr( iEndVal1 + 1 ).toLowerCase();
    }
WScript.Echo( szArg.toLowerCase() );
    switch( szArg.toLowerCase() )
    {
        case "/?":
            return false;

        case "/array":
            return CheckArrayCmd( szValue );

        case "/css":
            return CheckCssName( szValue );
            
        case "/cssDomain":
            return CheckCssDomain( szValue );
            
        case "/cssUser":
            return CheckCssUser( szValue );
            
        case "/cssPass":
            return CheckCssPass( szValue );
            
        case "/debug":
            return CheckDebugCmd( );

        case "/enterprise":
            return CheckEntCmd( );

        case "/logpath":
            return CheckLogPathCmd( szValue );

        case "/quiet":
            return CheckQuietCmd( );

        case "/server":
            return CheckServerCmd( szValue );

        case "/serveronly":
            return CheckServerOnlyCmd( );

        case "/tier":
            return CheckTierCmd( szValue );

        default:
            LogMessage( g_oMessages.L_InvalidOpt_txt + szArg + "\"" );
            return false;
    }
}

/**********************************************************************
 * CheckArrayCmd( szValue )
 * This function:
 *    1. Evaluates the "/array" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckArrayCmd( szValue )
{
    EnterFunction( arguments );

    if( "" != g_oVariables.szIsaArray )
    {
        return false;
    }
    g_oVariables.szIsaArray = szValue;
    return true;
}

/**********************************************************************
 * CheckCssName( szValue )
 * This function:
 *    1. Evaluates the "/array" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckCssName( szValue )
{
    EnterFunction( arguments );

    if( "" != g_oVariables.szCssName )
    {
        return false;
    }
    g_oVariables.szCssName = szValue;
    return true;
}

/**********************************************************************
 * CheckCssDomain( szValue )
 * This function:
 *    1. Evaluates the "/array" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckCssDomain( szValue )
{
    EnterFunction( arguments );

    if( "" != g_oVariables.szCssDomain )
    {
        return false;
    }
    g_oVariables.szCssDomain = szValue;
    return true;
}

/**********************************************************************
 * CheckCssUser( szValue )
 * This function:
 *    1. Evaluates the "/array" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckCssUser( szValue )
{
    EnterFunction( arguments );

    if( "" != g_oVariables.szCssUser )
    {
        return false;
    }
    var oSlash = /(\\ | \/)/g;
    var szDomain = "";
    if( oSlash.test( szValue ) )
    {
        var iSlash = szValue.indexOf( "\\" );
        if( 0 > iSlash )
        {
            iSlash = szValue.indexOf( "/" );
        }
        if( "" != g_oVariables.szCssDomain )
        {
            g_oVariables.szCssDomain = szValue.substr( 0, iSlash );
        }
        g_oVariables.szCssUser = szValue.substr( iSlash + 1 );
    }
    else
    {
        g_oVariables.szCssUser = szValue;
    }
    return true;
}

/**********************************************************************
 * CheckCssPass( szValue )
 * This function:
 *    1. Evaluates the "/array" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckCssPass( szValue )
{
    EnterFunction( arguments );

    if( "" != g_oVariables.szCssPass )
    {
        return false;
    }
    g_oVariables.szCssPass = szValue;
    return true;
}

/**********************************************************************
 * CheckDebugCmd( szValue )
 * This function:
 *    1. Evaluates the "/debug" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckDebugCmd( )
{
    EnterFunction( arguments );

    g_oVariables.fDebugMode = true;
    return true;
}

/**********************************************************************
 * CheckEntCmd( szValue )
 * This function:
 *    1. Evaluates the "/Ent" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckEntCmd( )
{
    EnterFunction( arguments );

    g_oVariables.fEntMode = true;
    return true;
}

/**********************************************************************
 * CheckLogPathCmd( szValue )
 * This function:
 *    1. Evaluates the "/logpath" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckLogPathCmd( szValue )
{
    EnterFunction( arguments );

    if( g_oVariables.fPathSet ||
        szValue == "" ||
        !g_oObjects.oFSO.FolderExists( szValue ) 
        )
    {
        LogMessage( 
            g_oMessages.L_NoFolder1_txt +
            szValue +  
            g_oMessages.L_NoFolder2_txt );
        return false;
    }
    if( szValue.lastIndexOf( "\\" ) != szValue.length - 1 )
    {
        szValue += "\\";
    }
    g_oVariables.szISAInfoPath = szValue;
    g_oVariables.fPathSet = true;
    return true;
}

/**********************************************************************
 * CheckQuietCmd( szValue )
 * This function:
 *    1. Evaluates the "/quiet" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckQuietCmd( )
{
    EnterFunction( arguments );

    if( g_oVariables.fQuiet == true )
    {
        return false;
    }
    g_oVariables.fQuiet = true;
    return true;
}

/**********************************************************************
 * CheckServerCmd( szValue )
 * This function:
 *    1. Evaluates the "/Server" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckServerCmd( szValue )
{
    EnterFunction( arguments );

    if( g_oVariables.fOneServer == true )
    {
        return false;
    }
    g_oVariables.szIsaServer = szValue;
    g_oVariables.fOneServer = true;
    return true;
}

/**********************************************************************
 * CheckServerOnlyCmd( szValue )
 * This function:
 *    1. Evaluates the "/ServerOnly" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckServerOnlyCmd( )
{
    EnterFunction( arguments );

    if( g_oVariables.fServerOnly == true )
    {
        return false;
    }
    g_oVariables.fServerOnly = true;
    return true;
}

/**********************************************************************
 * CheckTierCmd( szValue )
 * This function:
 *    1. Evaluates the "/Tier" cmd-line option
 *        
 *    2. calls into
 *        - none -
 *  3. called by 
 *        CheckArgument()
 *
 *    errors are not evaluated
 *********************************************************************/
function CheckTierCmd( szValue )
{
    EnterFunction( arguments );

    var iValue = parseInt( szValue );
    if( g_oVariables.fTier == true ||
        isNaN( iValue ) ||
        ( iValue < 0 || iValue > 3 ) )
    {
        return false;
    }
    g_oVariables.iTier = iValue;
    g_oVariables.fTier = true;
    return true;
}

/**********************************************************************
 * ShowUsage( )
 * This function:
 *    1. Displays the usage syntax
 *        
 *    2. calls into
 *        LogMessage
 *  3. called by 
 *        - nearly all functions -
 *
 *    errors are not evaluated
 *********************************************************************/
function ShowUsage( )
{
    EnterFunction( arguments );

    /*
     * see if we were asked for help
     */
    if( g_oMessages.szCmdOpts.indexOf( "?" ) == -1 )
    {
        WScript.Echo( g_oMessages.L_BadCommand_txt );
    }
    WScript.Echo( g_oMessages.L_Usage_txt );
}


/**********************************************************************
 * AskForYN( )
 * This function:
 *    1. Pops up a dialog with user defined title and prompt.
 *  2. Yes and No buttons
 *
 * if the Yes button is clicked:
 *     1. returns true
 *
 * if the No button is clicked:
 *    1. returns false
 *********************************************************************/
function AskForYN( szPrompt )
{
    EnterFunction( arguments );

    var iAnswer;
    var iTimeOut = 0;
    var iIcon = 32;
    var iYesNo = 4; /* popup dialog has "Yes" and "No" buttons */
    var iYes = 6; /* "Yes" button was clicked */
    var iDefault = 256    /* we want "no" to be the default answer */

    iAnswer = g_oObjects.oWsh.Popup( szPrompt, iTimeOut,
                            g_oMessages.L_TitleMsg_txt,
                            iYesNo + iIcon + iDefault );

    return ( iAnswer == iYes );
}

/*#######################################
 # End of helper functions
 ######################################*/

/*#######################################
    Debugging Functions
 ######################################*/
/**********************************************************************
 EnterFunction( oArguments )
    1. reports entry into function and arguments passed
    2. calls into
        GetFunctionData()
        Fprintf()
    3. called by 
        - many -
    4. Accepts
        function arguments object
    5. returns
        - nothing -
    6. outputs
        function data as string
**********************************************************************/
function EnterFunction( oArguments )
{
    if( g_oVariables.fDebugMode )
    {
        var arrData = GetFunctionData( oArguments );
        var szMsg = Fprintf( "  -->  %1(%2) from %3().", arrData );
        LogMessage( szMsg );
    }
}


/**********************************************************************
 ExitFunction( oArguments, vValue )
    1. reports exit from function and value returned
    2. calls into
        GetObjectData()
        GetFunctionData()
        Fprintf()
    3. called by 
        - many -
    4. Accepts
        function arguments object, "return" value
    5. returns
        - nothing -
    6. outputs
        function data as string
**********************************************************************/
function ExitFunction( oArguments, vValue )
{
    if( g_oVariables.fDebugMode )
    {
        var szMsg = "  <--  %1(%2) returned %4 to %3().";
        var arrData = GetFunctionData( oArguments );
        arrData.push( GetObjectData( vValue ) );
        szMsg = Fprintf( szMsg, arrData );
        LogMessage( szMsg );
    }
}



/**********************************************************************
 GetFunctionData( oArguments )
    1. obtains details of related function from arguments object
    2. calls into
        GetFunctionName()
        GetFunctionArgs()
    3. called by 
        EnterFunction()
        ExitFunction()
    4. Accepts
        function arguments object
    5. returns
        array of CalleeFunctionName, FunctionArgs, CallerFunctionName
        strings
    6. outputs
        - nothing -
**********************************************************************/
function GetFunctionData( oArguments )
{
    var szCallee = GetFunctionName( oArguments.callee );
    var szCaller = GetFunctionName( eval( szCallee + ".caller" ) );
    var arrData = new Array( szCallee,
                GetFunctionArgs( oArguments ),
                szCaller
                );
    return arrData;
}



/**********************************************************************
 GetFunctionName( oFunction )
    1. obtains function name from function object
    2. calls into
        - none -
    3. called by 
        GetFunctionData()
    4. Accepts
        function object
    5. returns
        function name as string
    6. outputs
        - nothing -
**********************************************************************/
function GetFunctionName( oFunction )
{
    if( !oFunction )
    {
        return "null";
    }
    var szFunction = oFunction.toString();
    var iStart = szFunction.indexOf( " " ) + 1;
    var iEnd = szFunction.indexOf( "(" );
    return szFunction.substr( iStart, iEnd - iStart );
}



/**********************************************************************
 GetFunctionArgs( oArguments )
    1. enumerates function arguments
    2. calls into
        GetObjectData()
        Fprintf()
    3. called by 
        GetFunctionData()
    4. Accepts
        function arguments object
    5. returns
        function arguments as string
    6. outputs
        - nothing -
**********************************************************************/
function GetFunctionArgs( oArguments )
{
    var arrArgs = new Array();
    for( var inx = 0; inx < oArguments.length; inx++ )
    {
        var vArg = oArguments[ inx ];
        var szData = GetObjectData( vArg );
        arrArgs.push( szData );
    }
    return arrArgs.join( ", " );
}



/**********************************************************************
 GetObjectData( oObject )
    1. obtains type and value of oObject
    2. calls into
        ObjToString()
        IsArray()
        GetArrayData()
    3. called by 
        GetFunctionArgs()
        GetArrayData()
    4. Accepts
        object
    5. returns
        array of object type and value strings
    6. outputs
        - nothing -
**********************************************************************/
function GetObjectData( oObject )
{
    var szType = "%1(%2)";
    var szTypeOf = typeof( oObject );
    var szValue = ObjToString( oObject ) ;

    if( IsArray( oObject ) )
    {
        szTypeOf = "array";
        szValue = GetArrayData( oObject );
    }
    return Fprintf( szType, new Array( szTypeOf, szValue ) );
}



/**********************************************************************
 GetArrayData( oArray )
    1. obtains type and value of array object members
    2. calls into
        GetObjectData()
        Fprintf()
    3. called by 
        GetFunctionArgs()
        GetArrayData()
    4. Accepts
        object
    5. returns
        array members type and value strings as string
    6. outputs
        - nothing -
**********************************************************************/
function GetArrayData( oArray )
{
    var arrData = new Array();
    for( var inx = 0; inx < oArray.length; inx++ )
    {
        arrData.push( GetObjectData( oArray[ inx ] ) );
    }
    return arrData.toString();
}



/**********************************************************************
 ObjToString( oObject )
    1. obtains string value of object
    2. calls into
        - nothing -
    3. called by 
        GetObjectData()
    4. Accepts
        object
    5. returns
        object value as string or empty string
    6. outputs
        - nothing -
**********************************************************************/
function ObjToString( oObject )
{
    //non-JScript objects don"t support toString()
    var szObject = "";
    try
    {
        szObject = oObject.toString();
    }
    catch( err )
    {
        szObject = "";
        err.clear;
    }
    return szObject;    
}



/**********************************************************************
 IsArray( oObject )
    1. determines if object is an array
    2. calls into
        - nothing -
    3. called by 
        GetObjectData()
    4. Accepts
        object
    5. returns
        true/false
    6. outputs
        - nothing -
**********************************************************************/
function IsArray( oObject )
{
    //non-JScript array objects don"t support join()
    try
    {
        oObject.join();
        return true;
    }
    catch( err )
    {
        err.clear;
        return false;
    }
}
/*#######################################
    End Of Debugging Functions
 ######################################*/

/*#######################################
 # Start of "Classes"
 ######################################*/

    /*#######################################
     # Start of objects
     ######################################*/
function IsaInfoObjects()
{
    this.oFSO                 = null;
    this.oHttpRequest       = null;
    this.oWsh            = null;
    this.oIsaXml            = null;        //ISA XML data object
    this.oISA                 = null;        //ISA COM root object
    this.oThisArray            = null;        //current ISA Array
    this.oThisServer        = null;        //current ISA server
    this.oIsaArray            = null;        //selected ISA Array
    this.oIsaServer            = null;        //selected ISA server
    this.fsTraceFile        = null;        //tracing log filestream object
    this.oWmiCimv2            = null;        //SWBem services object
    this.oWmiReg            = null;        //SWbem registry object
}
    /*#######################################
     # End of objects
     ######################################*/

    /*#######################################
     # Start of general variables
     ######################################*/
function IsaInfoVariables()
{
    this.lS_OK                    = 0;        //default "all ok" return value
    this.iCacheMode             = 4
    this.iFWMode                 = 59
    this.iIntegMode             = 63
    this.lHKLM                    = 0x80000002;
    this.iArray                    = 2;        //Ent Ed. in Array mode
    this.iMaxArgs                = 8;
    this.iTier                    = 1;        //scanning depth  -default is 1
    this.fDebugMode                = false;    //determines logging level
    this.fPathSet                = false;    //parser check for save path option
    this.fTier                    = false;    //parser check for scanning depth
    this.fIsa2k                 = false;    //ISA 2000 == true
    this.fSA                    = false;    //true if Isa2KEE in Standalone Mode
    this.fEE                    = false;    //true if EntEd
    this.fEntMode               = false;    //true if scanning the entire Enterprise
    this.fAdminOnly                = false;    //true if admin host
    this.fLocal_RRAS            = false;    //RRAS running state for local server
    this.fIsW2K3                = false;    //used for netstat variation
    this.fIsIIS                    = false;    //used to scan IIS bindings
    this.fIsDns                    = false;    //used to scan DNS bindings
    this.fSaveXml                = false;    //bool to trigger XML saves
    this.fMPSReports            = false;    //MPSReports compatibility flag
    this.fOneServer                = false;    //scan only this server
    this.fOneArray                = true;    //scan only the current array
    this.fServerOnly            = false;    //scan only this server; no export
    this.fQuiet                    = false;    //no msgboxes
    this.szComSpec                 = "";
    this.szSysFolder            = "";       //path to local system folder
    this.szISAInfoPath            = "";        //path for all ISAInfo files
    this.szXmlFile                = "";        //filename for the output XML
    this.szTempFilePath            = "";        //path for the temp files
    this.szThisServer            = "";        //placeholder for this computer name
    this.szThisArray            = "";        //placeholder for this array
    this.szIsaServer            = "";        //placeholder for the selected ISA server
    this.szIsaArray                = "";        //placeholder for the selected ISA Array
    this.szConfServer            = "";        //config strg server for 2k4ee
    this.szThisUser                = "";        //placeholder for interactive acct name
    this.szCssName                  = ""        //Configuration Storage Server
    this.szCssUser          = ""        //CSS user name
    this.szCssDomain        = ""        //CSS user domain
    this.szCssPass          = ""        //CSS User Password
    this.szErrNotFound             = "80070002";        //E_NOT_FOUND
    this.szErrNotSupported        = "800A01B6";    //method/property not supported
    this.szErrExists             = "800700B7";    //item already exists
    this.szUAlpha                = "\"ABCDEFGHIJKLMNOPQRSTUVWXYZ\"";
    this.szLAlpha                = "\"abcdefghijklmnopqrstuvwxyz\"";
    this.szIsaRegRoot            = "HKLM\\SOFTWARE\\Microsoft\\Fpc\\";
    this.szIsaEdition            = this.szIsaRegRoot + "Edition";
    this.szAcmeComp                = this.szIsaRegRoot + "AcmeComonents";
    this.szArrayGuid            = this.szIsaRegRoot + "CurrentArrayGuid";
    this.szSEGuid                = "{5933b383-0f51-46de-a54c-e9838062ecaf}";
    this.szEEGuid                = "{563f7924-6ac8-4ae6-bbfb-189dca3dd594}";
    this.szValidArgs            = new Array( "arr", "svr" );
    this.vFilter                 = /\%s/i;    //"%s" is the insertion trigger for Fprintf
}
    /*#######################################
     # End of general variables
     ######################################*/

    /*#######################################
     # Start of Localizable data
     ######################################*/
function IsaInfoMessages()
{
    this.szDivider             = "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+\r\n";
    this.L_Ver_txt                = "1.0.2161.23";            //current version
    this.szCmdOpts                = "";
    this.szScriptEngine            = "";
    this.szScriptFileName        = WScript.ScriptName;
    this.szScriptName            = this.szScriptFileName.substr( 0, this.szScriptFileName.indexOf( "." ) );
    this.L_TitleMsg_txt            = this.szScriptName + " version " + this.L_Ver_txt;
    this.szHeaderMsg            = "";
    this.L_Comment1_txt            = "ISA Server configuration exported by ";
    this.L_Comment2_txt            = " using " + this.L_TitleMsg_txt + " on " + 
                                    new Date().toString();
    this.szThisServer            = "";
                            
    /*    
     ************  error messages  ***************
     */
    this.L_Isa2k4Only_txt        = "\t*** This script only functions on ISA 2004 or 2006 servers.";
    this.L_ObjCreateErr_txt        = "\t*** Error encountered while creating ";
    this.L_SvrNotLocal_txt      = "%1 is not the local host; WMI queries are not possible to remote ISA servers.";
    this.L_NoFolder1_txt        = "\t*** Folder \"";
    this.L_NoFolder2_txt        = "\" couldn\"t be found.";
    this.L_CreateFldr_txt        = "\r\n\r\nWould you like to create it (Y/N)?";
    this.L_BadPath_txt            = " is not a valid path.";
    this.L_InvalidOpt_txt        = "\t*** Passed an invalid option : \"";
    this.L_NoWsh_txt        = "\t*** Unable to access the WScript.Shell object";
    this.L_NoFso_txt            = "\t*** Unable to access the FileScripting object";
    this.L_NoWinHttp_txt    = "\t*** Unable to access te WinHttp object.";
    this.L_NoLogFile_txt        = "\t*** Unable to create ";
    this.L_NoXmlSave_txt        = "\t*** Unable to save ";
    this.L_NoRras_txt            = "\t*** Unable to access the RemoteAccess Service.";
    this.L_NoWmiCimv2_txt        = "\t*** Unable to access the root\\default WMI namespace on ";
    this.L_NoWmiReg_txt            = "\t*** Unable to access the StdRegProv WMI class on ";
    this.L_SvrConnErr_txt        = "\t*** Failed to connect to WMI on ";
    this.L_QueryFailed_txt        = "\t*** Failed to query ";
    this.L_NoIsaData_txt        = "\t*** Unable to access MSXML2.DomDocument3.0 on ";
    this.L_GetISAFailed_txt        = "\t*** Unable to access the ISA Admin COM object.";
    this.L_SetupFailed_txt        = "\t*** Failed to complete the setup function tasks.";
    this.L_IsaDataFailed_txt     = "\t*** Failed to gather ISA configuration data.";
    this.L_GetSigAlertsFail_txt = "\t*** Failed to gather signaled alerts from ";
    this.L_GetSvrFailed_txt        = "\t*** Failed to gather Server data.";
    this.L_CloseFailed_txt        = "\t*** Failed to close ";
    this.L_SaveFailed_txt        = "\t*** Failed to save ";
    this.L_ExportFailed_txt        = "\t*** Failed to export ISA 2004 configuration data.";
    this.L_RunFailed_txt        = "\t*** Failed to execute ";
    this.L_GetFileFailed_txt    = "\t*** Failed to access ";
    this.L_EnumKeysFail_txt        = "\t*** Failed to enumerate registry keys in ";
    this.L_EnumValuesFail_txt     = "\t*** Failed to enumerate registry values in ";
    this.L_ReadFailed_txt        = "\t*** Failed to read ";
    this.L_ParseFailed_txt        = "\t*** Failed to parse g_oObjects.oIsaXml.";
    this.L_ConvertFailed_txt    = "\t*** Failed to convert g_oObjects.oIsaXml.xml.";
    this.L_ExtRunFailed_txt        = " returned %ErrorLevel% = ";
    this.L_CreateFailed1_txt    = "\t*** Failed to create a ";
    this.L_CreateFailed2_txt    = " child node with \"";
    this.L_CreateFailed3_txt    = "\" ";
    this.L_AccessFailed_txt        = "\t*** Failed to access ";
    this.L_LogFileSecFailed_txt = "Failed to access LogicalShareSecuritySetting for ";
    this.L_ErrNum_txt            = "\r\n\r\nError Number : ";
    this.L_ErrDesc_txt            = "\r\nDescription  : ";
    this.L_ErrSource_txt        = "\r\nSource       : ";
    this.L_Err_Line_txt            = "\r\nLine         : ";
    this.L_Err_Char_txt            = "\r\nChar         : ";
    this.L_Err_Text_txt            = "\r\nText         : ";
    this.L_CopyMsg_txt            = "\r\n\r\nHit <Ctrl>-C to copy this message to the clipboard.";
    this.L_Continue_txt            = "\r\n\r\nDo you wish to continue (\"No\" cancels)?";
    this.L_Quitting_txt            = this.szScriptName + " cannot continue...";
    this.L_Canceled_txt            = "\r\n\r\n" + this.szScriptName + " run canceled by user.";

    /*    
     ************  logging messages  ***************
     */
    this.L_RunningOn_txt        = "Running on ";
    this.L_RunningAs_txt        = " as ";
    this.L_Start_txt            = "Started ";
    this.L_AccessObj_txt        = " -- Accessing \"";
    this.L_CreateLog_txt        = " -- Creating ";
    this.L_SaveFile_txt            = " -- Saving ";
    this.L_AccessWMI_txt        = " -- Accessing WMI on ";
    this.L_GetISA_txt            = " -- Accessing ISA Admin COM...";
    this.L_ReadIsaData_txt        = " -- Exporting ISA configuration (this may take some time)...";
    this.L_IsaSvrData_txt        = " -- Reading additional ISA Server data...";
    this.L_CheckRras_txt        = " -- Determining RemoteAccess service state...";
    this.L_GetSvrs_txt            = " -- Enumerating the servers in this array...";
    this.L_GetEvtLogs_txt        = " -- Reading the event logs...";
    this.L_SigAlerts_txt        = " -- Reading ISA signaled alerts...";
    this.L_GetShares_txt        = " -- Enumerating the shares on this server...";
    this.L_GetIIS_txt            = " -- Examining the IIS service on this computer...";
    this.L_FoundIIS_txt            = " IIS services are installed on this computer...";
    this.L_RunCmd_txt            = " -- Executing ";
    this.L_GetFile_txt            = " -- Reading ";
    this.L_WmiQuery_txt            = " -- Executing WMI query \"";
    this.L_QueryResults_txt        = " items were returned from the query.";
    this.L_QueryLimited_txt      = "Event log queries cannot be counted.";
    this.L_EnumKey_txt            = " -- Reading registry tree at ";
    this.L_Cleanup_txt            = " -- Cleaning up the XML...";
    this.L_LogConfig_txt        = " -- Reading Array Logging configuration.";
    this.L_AllDone_txt            = "\r\n\r\n" + this.L_TitleMsg_txt + " completed at ";
    this.L_BadCommand_txt        = "\r\nSorry, one or more of your specified options was not valid.\r\n\t";
    this.L_SvrNotLocal_txt        = " -- %s is not the local machine - please run ISAInfo on that machine" +
                                  "\r\n with the /serveronly option to obtain detailed information from that server.";
    this.L_Usage_txt            = "\r\n" + this.szDivider + this.szDivider +
"\r\n" +
"     " + this.szScriptName + " is intended to provide an automated mechanism for ISA Server\r\n" +
"     2004 administrators to report their machine and array configuration to \r\n" +
"     support PSS and self-help troubleshooting.\r\n" +
"\r\n" +
"     " + this.szScriptName + " will scan ISA array and server configuration data and save \r\n" +
"     this data to an XML file and the script actions to a log file.  By \r\n" +
"     default, these files are located on the user\"s desktop.\r\n" +
"\r\n" +
"     If no command-line options are used, szScriptName will attempt to scan \r\n" +
"     all available servers in all available arrays.\r\n" +
"\r\n" +
"     " + this.szScriptName + " can be run with the following command line options:\r\n\r\n" +
"         cscript " + this.szScriptFileName + " [/?] [/array] [/debug] [/logpath] [/server]\r\n" +
"\r\n" +
"     All options are described below.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/?]\r\n\r\n" +
"     The /? option prints this help text.  For instance:\r\n" +
"\r\n" +
"         cscript " + this.szScriptName + " /?\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/array]\r\n\r\n" +
"     The /array option can be optionally followed by a name to specify the \r\n" +
"     desired ISA array to be scanned.  For instance:\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /array:ThisArray \r\n" +
"     will limit the configuration scan to the \"ThisArray\" array.\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /array\r\n" +
"     by itself will only scan the array where the script is running.\r\n" +
"\r\n" +
"\r\n" +
"     If this option is omitted, " + this.szScriptName + " will scan all available arrays.\r\n" +
"\r\n" +
"     If the specified array cannot be found, " + this.szScriptName + " will report an error\r\n" +
"     and exit.\r\n" +
"\r\n" +
"     For ISA Server Standard Edition, this option will not change " + this.szScriptName + "\r\n" +
"     behavior\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/css]\r\n\r\n" +
"     The /css option specifies the name of the CSS.\r\n" +
"\r\n" + 
"     This option is only required if " + this.szScriptName + " is executed from an admin-only" +
"\r\n" +
"     installation and requires the use of the CssDomain, CssName and CssPass " +"\r\n" +
"     options.  For instance:\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /css:CssName \r\n" +
"     will cause " + this.szScriptName + " to attempt a connection to the CSS at \'CssName\'.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/cssdomain]\r\n\r\n" +
"     The /cssdomain option specifies the domain context for the CSS user account.\r\n" +
"\r\n" + 
"     This option is only required if " + this.szScriptName + " is executed from an admin-only" +
"\r\n" +
"     installation.  For instance:\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /cssdomain:CssDomain \r\n" +
"     will cause " + this.szScriptName + " to authenticate to the CSS as a user from \'CssDomain\'.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/cssuser]\r\n\r\n" +
"     The /cssuser option specifies the CSS user account.\r\n" +
"\r\n" + 
"     This option is only required if " + this.szScriptName + " is executed from an admin-only" +
"\r\n" +
"     installation.  For instance:\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /cssuser:CssUser \r\n" +
"     will cause " + this.szScriptName + " to authenticate to the CSS as \'CssUser\'.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/csspass]\r\n\r\n" +
"     The /csspass option specifies the CSS user account password.\r\n" +
"\r\n" + 
"     This option is only required if " + this.szScriptName + " is executed from an admin-only" +
"\r\n" +
"     installation.  For instance:\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + " /csspass:CssPass \r\n" +
"     will cause " + this.szScriptName + " to use \'CssPass\' as the password when authenticating\r\n" +
"     as \'CssUser\'.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/debug]\r\n\r\n" +
"     The /debug option will increase the logging level so that any problems\r\n" +
"     encountered while running " + this.szScriptName + "  can be more easily identified and\r\n" +
"     corrected.  By default, this log is saved to the interactive user\"s \r\n" +
"     desktop, but using the /logpath option can change this.\r\n" +
"\r\n" +
"     This option functions best when it is used as the first option.\r\n" +
"\r\n" + this.szDivider + 
"\r\n" +
"        [/enterprise]\r\n\r\n" +
"     The /enterprise option will cause " + this.szScriptName + " to acquire the entire policy set.\r\n" +
"\r\n" + 
"     This action can take a *very* long time, depending on the size of the" +
"\r\n" +
"     Enterprise definition and should be used with care." +
"\r\n" +
"        cscript " + this.szScriptFileName + " /enterprise\r\n" +
"\r\n" + this.szDivider +
"\r\n" +
"        [/logpath:FullPathToFolder]\r\n\r\n" +
"     The /logpath option must be followed by a valid folder path.  UNC paths \r\n" +
"     are allowed.  The path nust be \"writable\" by the logged-in user.\r\n" +
"\r\n" +
"     You may not specify the ouput filename, which will always be:\r\n" +
"        " + this.szScriptName + "_ComputerName.xml\r\n" +
"     and \r\n" +
"        " + this.szScriptName + "_ComputerName.log\r\n" +
"\r\n" +
"     For instance:\r\n" +
"        cscript " + this.szScriptFileName+ "  /logpath:C:\\ISAInfo\r\n" +
"     will save the output files to \r\n" +
"        c:\\ISAInfo\\" + this.szScriptName + "_ComputerName.xml\r\n" +
"     and\r\n" +
"        c:\\ISAInfo\\" + this.szScriptName + "_ComputerName.log\r\n" +
"\r\n" +
"     The default path for " + this.szScriptName + " logging is the current desktop.\r\n" +
"\r\n" +
"     If the specified path cannot be found, " + this.szScriptName + " will report an error\r\n" +
"     and exit.\r\n" +
"\r\n" + this.szDivider +
"\r\n" +
"        [/quiet]\r\n\r\n" +
"     The /quiet option disables error message popups for unattended \r\n" +
"     execution.\r\n" +
"\r\n" +
"     For instance:\r\n" +
"        cscript " + this.szScriptFileName + "  /quiet\r\n" +
"     will not prompt the user if an error is encountered.\r\n" + 
"     Normally, the user is prompted to decide if the script should continue.\r\n" +
"\r\n" + this.szDivider +
"\r\n" +
"        [/server]\r\n\r\n" +
"     The /server option can be optionally followed by a name to limit the \r\n" +
"     server scanning to a specific machine.\r\n" +
"\r\n" +
"     For instance:\r\n" +
"        cscript " + this.szScriptFileName + "  /server:ComputerName \r\n" +
"     will limit the configuration scan to \"ComputerName\" only.\r\n" +
"\r\n" +
"        cscript " + this.szScriptFileName + "  /server\r\n" +
"     will only scan the server where the script is running\r\n" +
"\r\n" +
"     If this option is omitted, " + this.szScriptName + " will scan all available servers.\r\n" +
"\r\n" +
"     If the specified server cannot be found, " + this.szScriptName + " will report an error\r\n" +
"     and exit.\r\n" +
"\r\n" +
"     For ISA Server Standard Edition, this option will not change " + this.szScriptName + "\r\n" +
"     behavior\r\n" +
"\r\n" + this.szDivider +
"\r\n" +
"        [/serveronly]\r\n\r\n" +
"     The /serveronly option limits the data scan to server-specific \r\n" +
"     information only.\r\n" +
"\r\n" +
"     For instance:\r\n" +
"        cscript " + this.szScriptFileName + "  /serveronly\r\n" +
"     will limit the configuration scan to the local machine data only.\r\n" +
"\r\n" +
"     This option will override the \"/array\" and \"/server\" options.\r\n" +
"\r\n" + this.szDivider + this.szDivider;
}
    /*#######################################
     # End of Localizable data
     ######################################*/

/*#######################################
 # End of "Classes"
 ######################################*/

/*#######################################
 # End of ISAInfo.js
 ######################################*/

