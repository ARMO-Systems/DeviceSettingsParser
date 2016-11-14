using System.Reflection;

namespace ArmoSystems.ArmoGet.DeviceSettingsParser
{
    internal static class Common
    {
        public const string IdentificationModeEnumName = "EIdentificationMode";
        public const string SdkTypeEnumName = "DeviceTypeInfo1XPO.SDKEType";

        public static string ConvertToShortType( string str )
        {
            switch ( str )
            {
                case "System.Int32":
                    return "int";
                case "System.String":
                    return "string";
                case "System.Boolean":
                    return "bool";
            }
            return str;
        }

        public static string GetValueFromProperty( PropertyInfo prop, Terminal terminal )
        {
            if ( prop.Name.Equals( "RiDefault" ) )
                return IdentificationModeEnumName + "." + terminal.RiDefault;
            if ( prop.Name.Equals( "FType" ) )
            {
                var intType = ( int ) prop.GetValue( terminal, null );
                return "DeviceTypeInfo1XPO.SDKEType." + GetSdkTypeById( intType );
            }
            switch ( prop.PropertyType.ToString() )
            {
                case "System.Boolean":
                    return prop.GetValue( terminal, null )?.ToString().ToLower();
                case "System.String":
                    return $"\"{prop.GetValue( terminal, null )}\"";
                default:
                    return prop.GetValue( terminal, null )?.ToString();
            }
        }

        private static string GetSdkTypeById( int id )
        {
            switch ( id )
            {
                case 1:
                    return "ZK_BW";
                case 2:
                    return "ZK_TFT";
                case 3:
                    return "ZK_iFace";
                case 4:
                    return "C3";
                case 5:
                    return "Smartec";
                default:
                    return "Unknown";
            }
        }
    }
}