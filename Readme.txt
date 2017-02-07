+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

     This package is Copyright (c) 2004 - 2006 Microsoft Corporation.  

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

     ISAInfo.hta (1.0.2161.25)
     Since ISA 2004 configuration data is XML-based, it's difficult for many 
     folks to read.  That's where this tool comes in.  It will allow you to 
     read an ISAInfo_.xml file or even any portion of an ISA export or 
     backup file.
     It will reformat the data so that it's easier to find the information 
     you want.   Search, cut & paste are all dependent on Internet Explorer 
     for the time being.   You MUST have MXSML.3.0 installed or it cannot 
     function (will error out).

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

     ISAInfo.js (1.0.2161.24)
     ISAInfo.js is intended to provide an automated mechanism for ISA Server
     2004 administrators to report their machine and array configuration to
     support PSS and self-help troubleshooting.

     ISAInfo will scan ISA array and server configuration data and save
     this data to an XML file and the script actions to a log file.  By
     default, these files are located on the user's desktop.

     If no command-line options are used, ISAInfo will attempt to scan
     all available servers in all available arrays.

     ISAInfo can be run with the following command line options:

        cscript ISAInfo.js [/?] [/array] [/debug] [/logpath] [/quiet] [/server] [/serveronly]

     All options are described below.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

        [/?]

     The /? option prints this help text.  For instance:

        cscript ISAInfo /?

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

        [/array]

     The /array option can be optionally followed by a name to specify the
     desired ISA array to be scanned.  For instance:

        cscript ISAInfo.js /array:ThisArray
     will limit the configuration scan to the "ThisArray" array.

        cscript ISAInfo.js /array
     by itself will only scan the array where the script is running.


     If this option is omitted, ISAInfo will scan all available arrays.

     If the specified array cannot be found, ISAInfo will report an error
     and exit.

     For ISA Server Standard Edition, this option will not change ISAInfo
     behavior

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

        [/debug]

     The /debug option will increase the logging level so that any problems
     encountered while running ISAInfo  can be more easily identified and
     corrected.  By default, this log is saved to the interactive user's
     desktop, but using the /logpath option can change this.

     This option functions best when it is used as the first option.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

        [/logpath:FullPathToFolder]

     The /logpath option must be followed by a valid folder path.  UNC paths
     are allowed.  The path nust be 'writable' by the logged-in user.

     You may not specify the ouput filename, which will always be:
        ISAInfo_ComputerName.xml
     and
        ISAInfo_ComputerName.log

     For instance:
        cscript ISAInfo.js  /logpath:C:\ISAInfo
     will save the output files to
        c:\ISAInfo\ISAInfo_ComputerName.xml
     and
        c:\ISAInfo\ISAInfo_ComputerName.log

     The default path for ISAInfo logging is the current desktop.

     If the specified path cannot be found, ISAInfo will report an error
     and exit.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    	[/quiet]
     The /quiet option disables error message popups for unattended 
     execution.

     For instance:
    	cscript  ISAInfo /quiet
     will not prompt the user if an error is encountered. 
     Normally, the user is prompted to decide if the script should continue.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    	[/server]
     The /server option can be optionally followed by a name to limit the 
     server scanning to a specific machine.

     For instance:
    	cscript  ISAInfo /server:ComputerName 
     will limit the configuration scan to "ComputerName" only.

    	cscript  ISAInfo /server
     will only scan the server where the script is running

     If this option is omitted,  ISAInfo will scan all available servers.

     If the specified server cannot be found,  ISAInfo will report an error
     and exit.

     For ISA Server Standard Edition, this option will not change  ISAInfo 
     behavior

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    	[/serveronly]
     The /serveronly option limits the data scan to server-specific 
     information only.

     For instance:
    	cscript  ISAInfo /serveronly
     will limit the configuration scan to the local machine data only.

     This option will override the /array and /server options.

+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
