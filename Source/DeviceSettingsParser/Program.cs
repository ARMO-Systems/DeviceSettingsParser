using System;
using System.Linq;

namespace ArmoSystems.ArmoGet.DeviceSettingsParser
{
    internal static class Program
    {
        [STAThread]
        private static int Main( string[] args )
        {
            if ( args.Count() < 2 )
                return -2;
            var deviceSettingsParser = new DeviceSettingsOperator( args[ 0 ] );

            var res = deviceSettingsParser.ManageFile();
            if ( !res )
                return -1;
            deviceSettingsParser.CreateFile( string.Format( @"{0}\{1}", args[ 1 ], @"Common\Database\TimexDatabaseDevices.cs" ) );
            deviceSettingsParser.CreateFileTerminalEnums( string.Format( @"{0}\{1}", args[ 1 ], @"Shared\XPO\Source\Devices\DeviceTypeInfo1Generated.cs" ) );
            deviceSettingsParser.CreateFileDeviceTypeInfo( string.Format( @"{0}\{1}", args[ 1 ], @"Shared\XPO\Source\Devices\DeviceTypeInfo1XPO.cs" ) );
            return 0;
        }
    }
}