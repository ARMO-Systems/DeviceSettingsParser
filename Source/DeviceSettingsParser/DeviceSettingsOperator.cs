//using Excel = Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ArmoSystems.ArmoGet.DeviceSettingsParser
{
    internal sealed class DeviceSettingsOperator
    {
        private const string Terminals = "'Меню Терминалы$'";
        private const string Punkti = "'Меню Пункты доступа$'";
        private const string Func = "Функции$";
        private const string RejimIden = "Режим идентификации$";
        private readonly List< string > fieldsWithAtributes;
        private readonly String fileName;
        private readonly PropertyInfo[] propInfo;
        private readonly List< Terminal > terminalList;
        private DataRow currentRow;
        private string funcs = string.Empty;
        private string shablon = string.Empty;

        public DeviceSettingsOperator( string file )
        {
            fileName = file;
            terminalList = new List< Terminal >();
            ConnectionString = GetconnectionString( true );
            var term = new Terminal();
            propInfo = term.GetType().GetProperties();
            fieldsWithAtributes = new List< string >();
        }

        private string ConnectionString { get; set; }

        public bool ManageFile()
        {
            try
            {
                if ( !string.IsNullOrEmpty( Terminals ) )
                    FillMenuTerminals( Terminals );
                else
                    return false;

                if ( !string.IsNullOrEmpty( Punkti ) )
                    FillMenuPunktiDostupa( Punkti );
                else
                    return false;

                if ( !string.IsNullOrEmpty( Func ) )
                    FillFunctions( Func );
                else
                    return false;

                if ( !string.IsNullOrEmpty( RejimIden ) )
                    FillRejimIden( RejimIden );
                else
                    return false;

                shablon = GetFuncSetupSettingsBody();
                funcs = GetSetupFuncs();
                return true;
            }
            catch ( Exception e )
            {
                MessageBox.Show( e.Message );
                return false;
            }
        }

        public void CreateFile( string fileN )
        {
            var file = new StreamWriter( fileN );
            file.Write( CreateFile() );
            file.Close();
        }

        public void CreateFileDeviceTypeInfo( string name )
        {
            var file = new StreamWriter( name );
            file.Write( CreateFileDeviceTypeInfo() );
            file.Close();
        }

        private string CreateFileDeviceTypeInfo()
        {
            var fileContent = new StringBuilder();
            fileContent.Append( @"using DevExpress.Xpo;

namespace ArmoSystems.Timex.Shared.XPO.Source.Devices
{
    public sealed partial class DeviceTypeInfo1XPO : XPObject
    {" );

            foreach ( var prop in propInfo.OrderBy( prop => prop.Name ) )
            {
                fileContent.Append( Environment.NewLine );
                var fieldName = prop.Name.Substring( 0, 1 ).ToLower() + prop.Name.Substring( 1, prop.Name.Length - 1 );
                fileContent.Append( string.Format( "        private {0} {1};", prop.Name == "RiDefault" ? "ERejimIdentifikacii" : ConvertToShortType( prop.PropertyType.ToString() ), fieldName ) );
            }

            foreach ( var prop in propInfo.OrderBy( prop => prop.Name ) )
            {
                fileContent.Append( Environment.NewLine );
                fileContent.Append( Environment.NewLine );
                var fieldName = prop.Name.Substring( 0, 1 ).ToLower() + prop.Name.Substring( 1, prop.Name.Length - 1 );

                AddAttributes( fileContent, prop.Name );
                fileContent.Append( string.Format( "        public {0} {1}", prop.Name == "RiDefault" ? "ERejimIdentifikacii" : ConvertToShortType( prop.PropertyType.ToString() ), prop.Name ) );
                fileContent.Append( Environment.NewLine );
                fileContent.Append( "        {" );
                fileContent.Append( Environment.NewLine );
                fileContent.Append( string.Format( "            {1}{0}{2}", fieldName, "get { return ", "; }" ) );
                fileContent.Append( Environment.NewLine );
                fileContent.Append( string.Format( "            {1}SetPropertyValue( {0}, ref {3}, value ){2}", "\"" + prop.Name + "\"", "set { ", "; }", fieldName ) );
                fileContent.Append( Environment.NewLine );
                fileContent.Append( "        }" );
            }

            fileContent.Append( Environment.NewLine + "    }" + Environment.NewLine + '}' );
            return fileContent.ToString();
        }

        private void AddAttributes( StringBuilder fileContent, string name )
        {
            if ( fieldsWithAtributes.Contains( name ) )
            {
                fileContent.Append( "        [IsReadOnlyIfVedomiy]" );
                fileContent.Append( Environment.NewLine );
            }

            switch ( name )
            {
                case "FAvailableAC":
                    fileContent.Append( "        [Persistent( \"FEth\" )]" );
                    fileContent.Append( Environment.NewLine );
                    break;
                case "FAvailableInFreeVersion":
                    fileContent.Append( "        [Persistent( \"FEtn\" )]" );
                    fileContent.Append( Environment.NewLine );
                    break;
                case "FAvailableTA":
                    fileContent.Append( "        [Persistent( \"FEtz\" )]" );
                    fileContent.Append( Environment.NewLine );
                    break;
            }
        }

        private string CreateFile()
        {
            var fileContent = new StringBuilder();
            fileContent.Append( @"using ArmoSystems.Timex.Shared.XPO.Source.Devices;
using ArmoSystems.Timex.Shared.XPO;
using DevExpress.Xpo;
namespace ArmoSystems.Timex.Common.Database" );
            fileContent.Append( Environment.NewLine );
            fileContent.Append( '{' );
            fileContent.Append( "public static partial class TimexDatabase" );
            fileContent.Append( Environment.NewLine );
            fileContent.Append( '{' );
            fileContent.Append( funcs );
            fileContent.Append( Environment.NewLine );
            fileContent.Append( shablon );
            fileContent.Append( '}' + Environment.NewLine + '}' );
            return fileContent.ToString();
        }

        public void CreateFileTerminalEnums( string fileN )
        {
            var fileContent = new StringBuilder();
            fileContent.Append( @"namespace ArmoSystems.Timex.Shared.XPO.Source.Devices" );
            fileContent.Append( Environment.NewLine );
            fileContent.Append( '{' );
            fileContent.Append( "public partial class DeviceTypeInfo1XPO" );
            fileContent.Append( Environment.NewLine );
            fileContent.Append( '{' );
            fileContent.Append( GetNameAndEnums() );
            fileContent.Append( '}' + Environment.NewLine + '}' );
            using ( var file = new StreamWriter( fileN ) )
                file.Write( fileContent.ToString() );
        }

        private static string GetEnumName( string termName )
        {
            return termName.Replace( "-", "_" ).Replace( " ", "__" );
        }

        private string GetNameAndEnums()
        {
            var sb = new StringBuilder();
            sb.AppendLine( "public enum eNames" );
            sb.AppendLine( "{" );
            terminalList.ForEach( term => sb.AppendLine( GetEnumName( term.Name ) + "," ) );
            sb.AppendLine( "Unknown" );
            sb.Remove( sb.Length - 2, 2 );
            sb.AppendLine( "}" );

            sb.AppendLine( "public static string GetName( eNames e ){switch ( e ){" );
            terminalList.ForEach( term => sb.AppendLine( "case eNames." + GetEnumName( term.Name ) + ": return \"" + term.Name + "\";" ) );
            sb.AppendLine( "} return string.Empty;" );
            sb.AppendLine( "}" );

            sb.AppendLine( "public static eNames GetEnumName( string name ){switch ( name ){" );
            terminalList.ForEach( term => sb.AppendLine( "case \"" + term.Name + "\": return eNames." + GetEnumName( term.Name ) + ";" ) );
            sb.AppendLine( "} return eNames.Unknown;" );
            sb.AppendLine( "}" );
            return sb.ToString();
        }

        private string GetSetupFuncs()
        {
            const string begin = "SetupDeviceTypeSettings(session,";
            var sb = new StringBuilder();
            sb.Append( "public static void CreateTypeSettings( Session session )" + Environment.NewLine + "{" );
            foreach ( var terminal in terminalList )
            {
                sb.Append( begin );
                for ( var i = 0; i < propInfo.Count(); i++ )
                {
                    var prop = typeof ( Terminal ).GetProperty( propInfo[ i ].Name );
                    sb.Append( @"/*" + propInfo[ i ].Name.ToLower() + @"*/" );
                    if ( propInfo[ i ].Name.Equals( "FType" ) )
                    {
                        var intType = ( int ) prop.GetValue( terminal, null );
                        sb.Append( "DeviceTypeInfo1XPO.SDKEType." + GetSDKTypeById( intType ) + ',' );
                        continue;
                    }
                    if ( propInfo[ i ].Name.Equals( "RiDefault" ) )
                    {
                        sb.Append( "ERejimIdentifikacii." + terminal.RiDefault + ',' );
                        continue;
                    }
                    switch ( prop.PropertyType.ToString() )
                    {
                        case "System.Boolean":
                            sb.Append( prop.GetValue( terminal, null ).ToString().ToLower() );
                            break;
                        case "System.String":
                            sb.Append( '"' + prop.GetValue( terminal, null ).ToString() + '"' );
                            break;
                        default:
                            sb.Append( prop.GetValue( terminal, null ) );
                            break;
                    }
                    if ( i != propInfo.Count() - 1 )
                        sb.Append( ',' );
                }
                sb.Append( ");" + Environment.NewLine );
            }
            sb.Append( Environment.NewLine + "}" );
            return sb.ToString();
        }

        private static string GetSDKTypeById( int id )
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

        private string GetFuncSetupSettingsBody()
        {
            const string funcNemt = "private static void SetupDeviceTypeSettings( Session session, ";
            var args = new StringBuilder();
            args.Append( funcNemt );
            for ( var i = 0; i < propInfo.Count(); i++ )
            {
                if ( propInfo[ i ].Name.Equals( "FType" ) )
                {
                    args.Append( "DeviceTypeInfo1XPO.SDKEType ftype," );
                    continue;
                }
                if ( propInfo[ i ].Name.Equals( "RiDefault" ) )
                {
                    args.Append( "ERejimIdentifikacii ridefault," );
                    continue;
                }
                args.Append( ConvertToShortType( propInfo[ i ].PropertyType.ToString() ) + " " + propInfo[ i ].Name.ToLower() );
                if ( i != propInfo.Count() - 1 )
                    args.Append( ',' );
            }
            args.Append( ')' + Environment.NewLine + '{' + Environment.NewLine );
            args.Append( "var typeSettingsXPO = DeviceTypeInfo1XPO.FindType( session, name ) ?? new DeviceTypeInfo1XPO( session ) { Name = name };" + Environment.NewLine );
            const string obj = "typeSettingsXPO.";
            for ( var i = 0; i < propInfo.Count(); i++ )
            {
                if ( propInfo[ i ].Name.Equals( "FType" ) )
                {
                    args.Append( obj + "SDKType=ftype" + ';' + Environment.NewLine );
                    continue;
                }

                args.Append( obj + propInfo[ i ].Name + '=' + propInfo[ i ].Name.ToLower() + ';' + Environment.NewLine );
            }
            args.Append( '}' );
            return args.ToString();
        }

        private void FillFunctions( string listName )
        {
            var dt = GetDataFromList( listName );
            foreach ( DataRow row in dt.Rows )
            {
                currentRow = row;

                var name = currentRow[ "Имя" ];
                if ( name == null )
                    continue;

                var term = terminalList.FirstOrDefault( t => t.Name.Equals( name ) );
                if ( term == null )
                    continue;

                //TODO оставить только функцию в Catch после выпуска 3.11
                try
                {
                    term.FType = GetIntegerFromCell( "Тип (1-ч/б, 2- цв#, 3-iface, 4 -С3)" );
                }
                catch ( ArgumentException )
                {
                    term.FType = GetIntegerFromCell( "Тип(1-ч/б, 2-цв#, 3-iface, 4-С3, 5-Smartec)" );
                }
                term.FMaxUsers = GetIntegerFromCell( "Пользователей" );
                term.FMaxOtpechatkov = GetIntegerFromCell( "ОП" );
                term.FMaxLico = GetIntegerFromCell( "Лицо" );
                term.FMaxVein = GetIntegerFromCell( "Вены" );
                term.FMaxCards = GetIntegerFromCell( "Карта" );
                term.FKod = GetIntegerFromCell( "Код" );
                term.FMaxRecords = GetIntegerFromCell( "События" );
                term.FEthernet = StringYesToBool( "Ethernet" );
                term.FRS232 = StringYesToBool( "RS232/RS485" );
                term.FUSBHost = StringYesToBool( "USB host" );
                term.FUSBClient = StringYesToBool( "USB client" );
                term.FWiegandOut = StringYesToBool( "Wiegand Out" );
                term.FWiegandIn = StringYesToBool( "Wiegand In" );
                term.FWebServer = StringYesToBool( "WEB сервер" );
                term.FVvodOtpechatkovIzTimex = StringYesToBool( "Использование для ввода ОП с ПО" );
                term.FMaxFuncButtons = GetIntegerFromCell( "F-кнопки" );
                term.FTochekRegistracii = GetIntegerFromCell( "Точки регистрации СУРВ" );
                term.FVivodPhoto = StringYesToBool( "Вывод фото" );
                term.FVivodImeni = StringYesToBool( "Вывод имени" );
                term.FZahvatPhoto = StringYesToBool( "Захват фото" );
                term.FGolosovieSoobsheniya = StringYesToBool( "Голос# cообщ#" );
                term.FSMS = StringYesToBool( "СМС" );
                term.FKodRabot = StringYesToBool( "Код работ" );
                term.FSignalSmeni = StringYesToBool( "Сигнал смены" );
                term.FLetoZima = StringYesToBool( "Переход Лето/Зима" );
                term.FSkud = StringYesToBool( "СКУД" );
                term.FPunktovDostupa = GetIntegerFromCell( "Пункты доступа" );
                term.FProx = GetIntegerFromCell( "Прокс# сч#" );
                term.FLocalniyZPP = StringYesToBool( "Локальный ЗПП" );
                term.FVedomiy = StringYesToBool( "Ведомый считыватель" );
                term.FIspSchKlav = StringYesToBool( "Исп# сч# с клав#" );
                term.FPrazdniki = StringYesToBool( "Праздники" );
                term.FVremennieZoni = GetIntegerFromCell( "Временные зоны" );
                term.FGruppiDostupa = StringYesToBool( "Группы доступа" );
                var twoValuesCell = GetTwoValuesFromCell( "ОП/код принуждения" );
                term.FFingerPrintPrinujdenie = twoValuesCell.Item1;
                term.FCodPrinuzhd = twoValuesCell.Item2;
                term.FZamok = StringYesToBool( "Замок" );
                term.FTrevojniyVihod = StringYesToBool( "Трев# Выход" );
                term.FAvailableInFreeVersion = StringYesToBool( "Бесплатная версия" );
                term.FAvailableAC = StringYesToBool( "Timex AC" );
                term.FAvailableTA = StringYesToBool( "Timex TA" );
                term.FAlgorithm9 = StringYesToBool( "Алгоритм 9" );
                term.FAlgorithm10 = StringYesToBool( "Алгоритм 10" );
                term.FShluz = StringYesToBool( "Шлюз" );
                term.FImportTemplates = StringYesToBool( "Импорт шаблонов" );
                term.FDopVhodov = GetIntegerFromCell( "Доп# входы" );
                term.FDopVihodov = GetIntegerFromCell( "Доп# выходы" );
                term.FIdentificationTypes = StringYesToBool( "Типы идентификации" );
                term.FUsingZpp = StringYesToBool( "Использование ЗПП" );
                term.FAccessToTerminal = StringYesToBool( "Доступ к терминалу" );
                term.FDisplayNameOnTerminal = StringYesToBool( "Название на терминале", "FDisplayNameOnTerminal" );
                term.FVeinPerEmp = GetIntegerFromCell( "Шаблонов вен на сотрудника" );
                term.FCustomWiegand = StringYesToBool( "Кастомизированный wiegand" );
            }
        }

        private void FillRejimIden( string listName )
        {
            var dt = GetDataFromList( listName );
            foreach ( DataRow row in dt.Rows )
            {
                currentRow = row;

                var name = currentRow[ "Имя" ];
                if ( name == null )
                    continue;

                var term = terminalList.FirstOrDefault( t => t.Name.Equals( name ) );
                if ( term == null )
                    continue;

                term.RiEnabled = StringYesToBool( "Доступен" );
                term.RiSupportModes = StringYesToBool( "Поддержка режимов" );
                term.RiFingerprintOrCodeOrCard = StringYesToBool( "ОП/КОД/КАРТА" );
                term.RiPin = StringYesToBool( "ПИН" );
                term.RiFingerprint = StringYesToBool( "ОП" );
                term.RiCard = StringYesToBool( "КАРТА" );
                term.RiCode = StringYesToBool( "КОД" );
                term.RiFingerprintOrCode = StringYesToBool( "ОП/КОД" );
                term.RiFingerprintOrCard = StringYesToBool( "ОП/КАРТА" );
                term.RiCodeOrCard = StringYesToBool( "КОД/КАРТА" );
                term.RiPinAndFingerprint = StringYesToBool( "ПИН&ОП" );
                term.RiFingerprintAndCode = StringYesToBool( "ОП&КОД" );
                term.RiFingerprintAndCard = StringYesToBool( "ОП&КАРТА" );
                term.RiCodeAndCard = StringYesToBool( "КОД&КАРТА" );
                term.RiFingerprintAndCodeAndCard = StringYesToBool( "ОП&КОД&КАРТА" );
                term.RiPinAndFingerprintAndCode = StringYesToBool( "ПИН&ОП&КОД" );
                term.RiFingerprintAndCardOrPin = StringYesToBool( "ОП&КАРТА/ПИН" );
                term.RiFace = StringYesToBool( "ЛИЦО" );
                term.RiFaceAndCode = StringYesToBool( "ЛИЦО&КОД" );
                term.RiFaceAndCard = StringYesToBool( "ЛИЦО&КАРТА" );
                term.RiFaceAndCodeAndCard = StringYesToBool( "ЛИЦО&КАРТА&КОД" );
                term.RiFaceOrCodeOrCard = StringYesToBool( "ЛИЦО/КОД/КАРТА" );
                term.RiVein = StringYesToBool( "ВЕНЫ" );
                term.RiVeinAndCard = StringYesToBool( "ВЕНЫ&КАРТА" );
                term.RiVeinAndCode = StringYesToBool( "ВЕНЫ&КОД" );
                term.RiVeinAndCardAndCode = StringYesToBool( "ВЕНЫ&КАРТА&КОД" );
                term.RiVeinOrCodeOrCard = StringYesToBool( "ВЕНЫ/КОД/КАРТА" );
                term.RiDefault = GetDefaultRejimIdentifikacii( dt );
            }
        }

        private void FillMenuPunktiDostupa( string listName )
        {
            var dt = GetDataFromList( listName );

            foreach ( DataRow row in dt.Rows )
            {
                currentRow = row;

                var name = currentRow[ "Имя" ];
                if ( name == null )
                    continue;
                var term = terminalList.FirstOrDefault( t => t.Name.Equals( name ) );
                if ( term == null )
                    continue;
                term.PdPunktiDostupa = StringYesToBool( "Пункты доступа", "PdPunktiDostupa" );
                term.PdSettings = StringYesToBool( "Настройки", "PdSettings" );
                term.PdSettingsDoorWorkByTimeZone = StringYesToBool( "Работа двери по временной зоне", "PdSettingsDoorWorkByTimeZone" );
                term.PdSettingsDoorUnlockByTimeZone = StringYesToBool( "Разблокировка двери по врем# зоне", "PdSettingsDoorUnlockByTimeZone" );
                term.PdSettingsDoorBlockByTimeZone = StringYesToBool( "Блокировка двери по врем# зоне", "PdSettingsDoorBlockByTimeZone" );
                term.PdSettingsDoorUnlockTimeout = StringYesToBool( "Время разблокировки замка (сек)", "PdSettingsDoorUnlockTimeout" );
                term.PdSettingsReadingDelay = StringYesToBool( "Задержка считывания (сек)", "PdSettingsReadingDelay" );
                term.PdSettingsIdentificationMode = StringYesToBool( "Режим идентификации", "PdSettingsIdentificationMode" );
                term.PdSettingsVedomiy = StringYesToBool( "Ведомый считыватель", "PdSettingsVedomiy" );
                term.PdDoorMonitoring = StringYesToBool( "Мониторинг двери", "PdDoorMonitoring" );
                term.PdDoorMonitoringSensorType = StringYesToBool( "Тип датчика", "PdDoorMonitoringSensorType" );
                term.PdDoorMonitoringCloseDoorBySensor = StringYesToBool( "Закрывать замок по датчику", "PdDoorMonitoringCloseDoorBySensor" );
                term.PdDoorMonitoringAlarmOpenDoor = StringYesToBool( "Тревога \"Дверь ост#откр\" через (сек)", "PdDoorMonitoringAlarmOpenDoor" );
                term.PdDoorMonitoringAlarmExitInterval = StringYesToBool( "Тревожный выход через (сек)", "PdDoorMonitoringAlarmExitInterval" );
                term.PdDopSchitivatel = StringYesToBool( "Дополнительный считыватель", "PdDopSchitivatel" );
                term.PdDopSchitivatelVedomiy = StringYesToBool( "Дополнительный ведомый", "PdDopSchitivatelVedomiy" );
                term.PdDopSchitivatelWorkMode = StringYesToBool( "Режим работы", "PdDopSchitivatelWorkMode" );
                term.PdDopSchitivatelName = StringYesToBool( "Название", "PdDopSchitivatelName" );
                term.PdExtra = StringYesToBool( "Дополнительно", "PdExtra" );
                term.PdExtraAccessCodeUnderForce = StringYesToBool( "Код доступа под принуждением", "PdExtraAccessCodeUnderForce" );
                term.PdExtraManualExitOnAlarmCount = StringYesToBool( "Общий выход по счетчику тревог", "PdExtraManualExitOnAlarmCount" );
                term.PdSettingsFirstCardOpenDoor = StringYesToBool( "Разблокировка по первой карте", "PdSettingsFirstCardOpenDoor" );
                term.PdExtraEmergencyCode = StringYesToBool( "Экстренный код" );
                term.PdExtraMultiCardOpenDoor = StringYesToBool( "Правило N лиц" );
                term.PdDoorOpen = StringYesToBool( "Открыть" );
                term.PdDoorUnblock = StringYesToBool( "Разблокировать" );
                term.PdDoorBlock = StringYesToBool( "Заблокировать" );
                term.PdVedomiyIdentificationMode = StringYesToBool( "Идентификация ведомого", "PdVedomiyIdentificationMode" );
                term.PdDopSchitivatelIdentificationMode = StringYesToBool( "Идентификация считывателя", "PdDopSchitivatelIdentificationMode" );
            }
        }

        private static string ConvertToShortType( string str )
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

        private Tuple< bool, bool > GetTwoValuesFromCell( string colName )
        {
            if ( !IsEnableColumn( colName ) || !( currentRow[ colName ] is string ) )
                return new Tuple< bool, bool >( false, false );

            var matches = new Regex( "(да|нет)+" ).Matches( currentRow[ colName ] as string );

            if ( matches.Count == 0 )
                return new Tuple< bool, bool >( false, false );
            if ( matches.Count != 1 )
                return new Tuple< bool, bool >( ObjectEqualsYes( matches[ 0 ] ), ObjectEqualsYes( matches[ 1 ] ) );
            var comBool = ObjectEqualsYes( matches[ 0 ] );
            return new Tuple< bool, bool >( comBool, comBool );
        }

        private static bool ObjectEqualsYes( object inObject )
        {
            return inObject.ToString().Equals( "да" );
        }

        private DataTable GetDataFromList( string listName )
        {
            var connection2 = new OleDbConnection( ConnectionString );

            var command = new OleDbCommand( "SELECT * FROM [" + listName + "]", connection2 );
            connection2.Open();
            var da = new OleDbDataAdapter( command );

            var dt = new DataTable();

            da.Fill( dt );
            return dt;
        }

        private void FillMenuTerminals( string listName )
        {
            var dt = GetDataFromList( listName );

            foreach ( DataRow row in dt.Rows )
            {
                currentRow = row;

                var term = new Terminal { Name = currentRow[ "Имя" ] as string };
                if ( ( term.Name == null ) || ( term.Name.EndsWith( "*", StringComparison.Ordinal ) ) )
                    continue;

                term.TcsConnectionConfigs = StringYesToBool( "Настройки соединения с терм#", "TcsConnectionConfigs" );
                term.TcsConnectionKey = StringYesToBool( "Ключ связи", "TcsConnectionKey" );
                term.TcsConnectionType = StringYesToBool( "Тип связи", "TcsConnectionType" );

                term.TcsEthernet = StringYesToBool( "Ethernet", "TcsEthernet" );
                term.TcsEthernetIpAddress = StringYesToBool( "IP адрес", "TcsEthernetIpAddress" );
                term.TcsEthernetIpAddressPort = StringYesToBool( "Порт", "TcsEthernetIpAddressPort" );

                term.TcsRs232Rs485 = StringYesToBool( "RS232/RS485", "TcsRs232Rs485" );
                term.TcsRs232Rs485Port = StringYesToBool( "Порт RS", "TcsRs232Rs485Port" );
                term.TcsRs232Rs485Speed = StringYesToBool( "Скорость предачи данных", "TcsRs232Rs485Speed" );
                term.TcsRs232Rs485DeviceAddress = StringYesToBool( "Адрес устройства", "TcsRs232Rs485DeviceAddress" );

                term.TsConfigs = StringYesToBool( "Настройки терминала", "TsConfigs" );
                term.TsEthernet = StringYesToBool( "Ethernet терминала", "TsEthernet" );
                term.TsEthernetIpAddress = StringYesToBool( "IP адрес терминала", "TsEthernetIpAddress" );
                term.TsEthernetMask = StringYesToBool( "Маска подсети", "TsEthernetMask" );
                term.TsEthernetGateway = StringYesToBool( "Шлюз", "TsEthernetGateway" );
                term.TsEthernetSpeed = StringYesToBool( "Скорость сети", "TsEthernetSpeed" );
                term.TsEthernetAllow = StringYesToBool( "Разрешить Ethernet", "TsEthernetAllow" );

                term.TsRs232Rs485 = StringYesToBool( "RS232/RS4851", "TsRs232Rs485" );
                term.TsRs232Rs485Speed = StringYesToBool( "Скорость предачи данных1", "TsRs232Rs485Speed" );
                term.TsRs232Rs485DeviceAddress = StringYesToBool( "Адрес устройства1", "TsRs232Rs485DeviceAddress" );
                term.TsRs232Allow = StringYesToBool( "Разрешить RS232", "TsRs232Allow" );
                term.TsRs485Allow = StringYesToBool( "Разрешить RS485", "TsRs485Allow" );

                term.TsTimeSync = StringYesToBool( "Синхронизация времени", "TsTimeSync" );

                term.TsTimeSyncTimeZone = StringYesToBool( "Часовой пояс", "TsTimeSyncTimeZone" );

                term.TsWiegandEnter = StringYesToBool( "Wiegand вход", "TsWiegandEnter" );
                term.TsWiegandEnterFormatType = StringYesToBool( "Тип Wiegand", "TsWiegandEnterFormatType" );
                term.TsWiegandEnterFormat = StringYesToBool( "Формат", "TsWiegandEnterFormat" );
                term.TsWiegandEnterBitsCount = StringYesToBool( "Число бит", "TsWiegandEnterBitsCount" );
                term.TsWiegandEnterImpulsDuration = StringYesToBool( "Длительность импульса", "TsWiegandEnterImpulsDuration" );
                term.TsWiegandEnterImpulsInterval = StringYesToBool( "Интервал импульса", "TsWiegandEnterImpulsInterval" );
                term.TsWiegandEnterDataType = StringYesToBool( "Тип данных", "TsWiegandEnterDataType" );

                term.TsWiegandExit = StringYesToBool( "Wiegand выход", "TsWiegandExit" );
                term.TsWiegandExitFormat = StringYesToBool( "Формат1", "TsWiegandExitFormat" );
                term.TsWiegandExitErrorCodeChecker = StringYesToBool( "Код ошибки (чек-бокс)", "TsWiegandExitErrorCodeChecker" );
                term.TsWiegandExitErrorCode = StringYesToBool( "Значение", "TsWiegandExitErrorCode" );
                term.TsWiegandExitObjectCodeChecker = StringYesToBool( "Код объекта (чек-бокс)", "TsWiegandExitObjectCodeChecker" );
                term.TsWiegandExitObjectCode = StringYesToBool( "Значение1", "TsWiegandExitObjectCode" );
                term.TsWiegandExitImpulsDuration = StringYesToBool( "Длительность импульса1", "TsWiegandExitImpulsDuration" );
                term.TsWiegandExitImpulsInterval = StringYesToBool( "Интервал импульса1", "TsWiegandExitImpulsInterval" );
                term.TsWiegandExitDataType = StringYesToBool( "Тип данных выхода", "TsWiegandExitDataType" );

                term.TsRecognition = StringYesToBool( "Распознавание", "TsRecognition" );
                term.TsRecognition1N = StringYesToBool( "Пороговый уровень 1:N", "TsRecognition1N" );
                term.TsRecognition11 = StringYesToBool( "Пороговый уровень 1:1", "TsRecognition11" );
                term.TsFaceRecognition1N = StringYesToBool( "Пороговый уровень 1:N (Лицо)", "TsFaceRecognition1N" );
                term.TsFaceRecognition11 = StringYesToBool( "Пороговый уровень 1:1 (Лицо)", "TsFaceRecognition11" );
                term.TsRecognition11Only = StringYesToBool( "Только 1:1", "TsRecognition11Only" );
                term.TsRecognitionCardOnly = StringYesToBool( "Только по карте", "TsRecognitionCardOnly" );

                term.TsEffecting = StringYesToBool( "Оформление", "TsEffecting" );
                term.TsEffectingDateFormat = StringYesToBool( "Формат даты", "TsEffectingDateFormat" );
                term.TsEffectingVoiceMessages = StringYesToBool( "Голосовые сообщения", "TsEffectingVoiceMessages" );
                term.TsEffectingVoiceMessageIdent = StringYesToBool( "Голосовая идентификация", "TsEffectingVoiceMessageIdent" );
                term.TsEffectingVolume = StringYesToBool( "Громкость", "TsEffectingVolume" );
                term.TsEffectingButtonsSound = StringYesToBool( "Звук кнопок", "TsEffectingButtonsSound" );
                term.TsEffectingPhotoOutput = StringYesToBool( "Вывод фото", "TsEffectingPhotoOutput" );
                term.TsEffectingNameOutput = StringYesToBool( "Вывод имени", "TsEffectingNameOutput" );

                term.TsOtherSettings = StringYesToBool( "Другие настройки", "TsOtherSettings" );
                term.TsOtherSettingsSposobnostPerehodaVRegimBezd = StringYesToBool( "Режим при бездействии", "TsOtherSettingsSposobnostPerehodaVRegimBezd" );
                term.TsOtherSettingsVremyaPerehodaVRegimBezd = StringYesToBool( "Таймаут при бездействии", "TsOtherSettingsVremyaPerehodaVRegimBezd" );
                term.TsOtherSettingsMenuExitTimeout = StringYesToBool( "Таймаут выхода из меню", "TsOtherSettingsMenuExitTimeout" );
                term.TsOtherSettingsPhotoCapture = StringYesToBool( "Захват фото", "TsOtherSettingsPhotoCapture" );
                term.TsOtherSettingsSMS = StringYesToBool( "СМС", "TsOtherSettingsSMS" );
                term.TsOtherSettingsWorksCode = StringYesToBool( "Код работ", "TsOtherSettingsWorksCode" );
                term.TsOtherSettingsChangeSignal = StringYesToBool( "Сигнал смены", "TsOtherSettingsChangeSignal" );
                term.TsOtherSettingsAutomaticRecognition = StringYesToBool( "Автоматическое распознавание", "TsOtherSettingsAutomaticRecognition" );

                term.TsUpperMenu = StringYesToBool( "Верхнее меню", "TsUpperMenu" );
                term.TsSyncTime = StringYesToBool( "Синхронизовать время", "TsSyncTime" );
                term.TsRestart = StringYesToBool( "Перезагрузить", "TsRestart" );
                term.TsShutdown = StringYesToBool( "Выключить", "TsShutdown" );
                term.TsRemoveAdmins = StringYesToBool( "Сброс админ# привелегий", "TsRemoveAdmins" );
                term.TsFlushAllData = StringYesToBool( "Сброс всех данных", "TsFlushAllData" );
                term.TsUpperMenuUpdateFirmware = StringYesToBool( "Обновление прошивки", "TsUpperMenuUpdateFirmware" );
                term.TsUpperMenuTimexUsb = StringYesToBool( "Timex <-> USB", "TsUpperMenuTimexUsb" );

                term.TsStatistic = StringYesToBool( "Статистика", "TsStatistic" );
                term.TsStatisticLastUpdateTime = StringYesToBool( "Время последнего обновления", "TsStatisticLastUpdateTime" );
                term.TsStatisticDeviceTime = StringYesToBool( "Время на устройстве", "TsStatisticDeviceTime" );
                term.TsStatisticUsers = GetIntegerFromCell( "Пользователи", "TsStatisticUsers" );
                term.TsStatisticFingerPrints = GetIntegerFromCell( "Шаблоны", "TsStatisticFingerPrints" );
                term.TsStatisticFace = GetIntegerFromCell( "Лицо", "TsStatisticFace" );
                term.TsStatisticCards = GetIntegerFromCell( "Карта", "TsStatisticCards" );
                term.TsStatisticCode = GetIntegerFromCell( "Код", "TsStatisticCode" );
                term.TsStatisticEvents = GetIntegerFromCell( "События", "TsStatisticEvents" );
                term.TsStatisticsAdmins = StringYesToBool( "Администраторы", "TsStatisticsAdmins" );
                term.TsStatisticSerialNumber = StringYesToBool( "Серийный номер", "TsStatisticSerialNumber" );
                term.TsStatisticSoft = StringYesToBool( "Прошивка", "TsStatisticSoft" );
                term.TsGroup = GetIntegerFromCell( "Группа устройства", "TsGroup" );

                term.comments = currentRow[ "Примечание" ] as string;

                terminalList.Add( term );
            }
        }

        private int GetIntegerFromCell( string colName )
        {
            int i;
            return IsEnableColumn( colName ) && int.TryParse( currentRow[ colName ].ToString(), out i ) ? i : 0;
        }

        private bool IsEnableColumn( string colName )
        {
            try
            {
                currentRow.Field< object >( colName );
                return true;
            }
            catch ( ArgumentException )
            {
                currentRow.Field< object >( string.Format( "{0}*", colName ) );
                return false;
            }
        }

        private bool StringYesToBool( string colName )
        {
            return IsEnableColumn( colName ) && ( currentRow[ colName ].ToString().Equals( "да" ) || currentRow[ colName ].ToString().Equals( "да*" ) );
        }

        private bool StringYesToBool( string colName, string property )
        {
            try
            {
                return StringYesToBool( colName );
            }
            catch ( ArgumentException )
            {
                var value = StringYesToBool( string.Format( "{0}$", colName ) );
                if ( !fieldsWithAtributes.Contains( property ) )
                    fieldsWithAtributes.Add( property );
                return value;
            }
        }

        private string GetDefaultRejimIdentifikacii( DataTable dt )
        {
            var colDefault = dt.Columns.Cast< DataColumn >().FirstOrDefault( col => IsDefaultBoolValue( col.ColumnName ) );
            return colDefault != null
                ? colDefault.ColumnName.Replace( "&", "And" ).
                    Replace( "/", "Or" ).
                    Replace( "ОП", "Fingerprint" ).
                    Replace( "КОД", "Code" ).
                    Replace( "ПИН", "Pin" ).
                    Replace( "КАРТА", "Card" ).
                    Replace( "ЛИЦО", "Face" ).
                    Replace( "ВЕНЫ", "Vein" )
                : String.Empty;
        }

        private bool IsDefaultBoolValue( string colName )
        {
            return currentRow[ colName ].ToString().Equals( "да*" );
        }

        private int GetIntegerFromCell( string colName, string property )
        {
            try
            {
                return GetIntegerFromCell( colName );
            }
            catch ( ArgumentException )
            {
                var value = GetIntegerFromCell( string.Format( "{0}$", colName ) );
                if ( !fieldsWithAtributes.Contains( property ) )
                    fieldsWithAtributes.Add( property );
                return value;
            }
        }

        private string GetconnectionString( bool is2007 )
        {
            return is2007
                ? @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\""
                : @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
        }
    }
}