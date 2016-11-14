﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ArmoSystems.ArmoGet.DeviceSettingsParser
{
    internal sealed class DevicesInfoCreator
    {
        private const string UsingCommonExternalSystems = "using ArmoSystems.Timex.Common.ExternalSystems;";
        private const string UsingSdkExternalSystems = "using ArmoSystems.Timex.SDK.ExternalSystem;";
        private const string UsingSourceDevices = "using ArmoSystems.Timex.Shared.XPO.Source.Devices;";
        private const string DeviceInfoNamespace = "namespace ArmoSystems.Timex.PullZKPlugin";
        private const string DeviceInfoInterfaceNamespace = "namespace ArmoSystems.Timex.Common.ExternalSystems";
        private const string DeviceInfoInterfaceStr = "IDeviceInfo";
        private const string Indent = "        ";
        private const string DirAutogeneratedDeviceInfo = @"D:\Temp\TimexTemp\Autogenerated\DevicesInfo";
        private readonly List< PropertyInfo > properties;
        private readonly List< Terminal > terminals;

        public DevicesInfoCreator( List< Terminal > terminals, List< PropertyInfo > properties )
        {
            this.terminals = terminals;
            this.properties = properties;
        }

        public void CreateAutogeneratedDevicesInfo()
        {
            using ( var file = new StreamWriter( $@"{DirAutogeneratedDeviceInfo}\{DeviceInfoInterfaceStr}.cs" ) )
                file.Write( CreateIDeviceInfo() );

            foreach ( var terminal in terminals )
            {
                var terminalClassName = terminal.Name.Replace( " ", "_" ).Replace( "-", "_" ).Replace( "ведомый", "slave" ).Replace( "дверь", "door" ).Replace( "двери", "doors" );
                using ( var file = new StreamWriter( $@"{DirAutogeneratedDeviceInfo}\{terminalClassName}.cs" ) )
                    file.Write( CreateDeviceInfo( terminal, terminalClassName ) );
            }
        }

        private static string GetPropertyType( PropertyInfo prop )
        {
            return prop.Name == "RiDefault" ? Common.IdentificationModeEnumName : prop.Name == "FType" ? Common.SdkTypeEnumName : Common.ConvertToShortType( prop.PropertyType.ToString() );
        }

        private string CreateDeviceInfo( Terminal terminal, string terminalClassName )
        {
            var fileContent = new StringBuilder();
            fileContent.Append( string.Join( Environment.NewLine, UsingCommonExternalSystems, UsingSdkExternalSystems, UsingSourceDevices, DeviceInfoNamespace, "{" ) );
            fileContent.Append( $"{Indent}public sealed class {terminalClassName}: {DeviceInfoInterfaceStr} {{" );
            fileContent.Append( string.Join( Environment.NewLine,
                properties.OrderBy( prop => prop.Name ).Select( prop => $"{Indent}public {GetPropertyType( prop )} {prop.Name} => {Common.GetValueFromProperty( prop, terminal )};" ) ) );
            fileContent.Append( string.Join( Environment.NewLine, "}", "}" ) );
            return fileContent.ToString();
        }

        private string CreateIDeviceInfo()
        {
            var fileContent = new StringBuilder();
            fileContent.Append( string.Join( Environment.NewLine, UsingSdkExternalSystems, UsingSourceDevices, DeviceInfoInterfaceNamespace, "{" ) );
            fileContent.Append( $"{Indent}public interface {DeviceInfoInterfaceStr} {{" );
            fileContent.Append( string.Join( Environment.NewLine, properties.OrderBy( prop => prop.Name ).Select( prop => $"{Indent}{GetPropertyType( prop )} {prop.Name} {{ get; }}" ) ) );
            fileContent.Append( string.Join( Environment.NewLine, "}", "}" ) );
            return fileContent.ToString();
        }
    }
}