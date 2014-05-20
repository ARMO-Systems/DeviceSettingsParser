namespace ArmoSystems.ArmoGet.DeviceSettingsParser
{
    internal class Terminal
    {
        public string comments;
        public string Name { get; set; }

        public bool TcsConnectionConfigs { get; set; }
        public bool TcsConnectionKey { get; set; }
        public bool TcsConnectionType { get; set; }

        public bool TcsEthernet { get; set; }
        public bool TcsEthernetIpAddress { get; set; }
        public bool TcsEthernetIpAddressPort { get; set; }
        public bool TcsRs232Rs485 { get; set; }
        public bool TcsRs232Rs485Port { get; set; }
        public bool TcsRs232Rs485Speed { get; set; }
        public bool TcsRs232Rs485DeviceAddress { get; set; }

        public bool TsConfigs { get; set; }
        public bool TsEthernet { get; set; }
        public bool TsEthernetIpAddress { get; set; }
        public bool TsEthernetMask { get; set; }
        public bool TsEthernetGateway { get; set; }
        public bool TsEthernetSpeed { get; set; }
        public bool TsEthernetAllow { get; set; }
        public bool TsRs232Rs485 { get; set; }
        public bool TsRs232Rs485Speed { get; set; }
        public bool TsRs232Rs485DeviceAddress { get; set; }
        public bool TsRs232Allow { get; set; }
        public bool TsRs485Allow { get; set; }

        public bool TsTimeSync { get; set; }
        public bool TsTimeSyncAuto { get; set; }
        public bool TsTimeSyncTimeZone { get; set; }
        public bool TsTimeSyncSummerWinter { get; set; }

        public bool TsWiegandEnter { get; set; }
        public bool TsWiegandEnterFormat { get; set; }
        public bool TsWiegandEnterBitsCount { get; set; }
        public bool TsWiegandEnterImpulsDuration { get; set; }
        public bool TsWiegandEnterImpulsInterval { get; set; }
        public bool TsWiegandEnterDataType { get; set; }
        public bool TsWiegandExit { get; set; }
        public bool TsWiegandExitFormat { get; set; }
        public bool TsWiegandExitErrorCode { get; set; }
        public bool TsWiegandExitErrorCodeChecker { get; set; }
        public bool TsWiegandExitObjectCode { get; set; }
        public bool TsWiegandExitObjectCodeChecker { get; set; }
        public bool TsWiegandExitImpulsDuration { get; set; }
        public bool TsWiegandExitImpulsInterval { get; set; }
        public bool TsWiegandExitDataType { get; set; }
        public bool TsRecognition { get; set; }
        public bool TsRecognition1N { get; set; }
        public bool TsRecognition11 { get; set; }
        public bool TsFaceRecognition1N { get; set; }
        public bool TsFaceRecognition11 { get; set; }
        public bool TsRecognition11Only { get; set; }
        public bool TsRecognitionCardOnly { get; set; }

        public bool TsEffecting { get; set; }
        public bool TsEffectingDateFormat { get; set; }
        public bool TsEffectingVoiceMessages { get; set; }
        public bool TsEffectingVoiceMessageIdent { get; set; }
        public bool TsEffectingVolume { get; set; }
        public bool TsEffectingButtonsSound { get; set; }
        public bool TsEffectingPhotoOutput { get; set; }
        public bool TsEffectingNameOutput { get; set; }

        public bool TsOtherSettings { get; set; }
        public bool FVedomiy { get; set; }
        public bool TsOtherSettingsSposobnostPerehodaVRegimBezd { get; set; }
        public bool TsOtherSettingsVremyaPerehodaVRegimBezd { get; set; }
        public bool TsOtherSettingsMenuExitTimeout { get; set; }
        public bool TsOtherSettingsPhotoCapture { get; set; }
        public bool TsOtherSettingsSMS { get; set; }
        public bool TsOtherSettingsWorksCode { get; set; }
        public bool TsOtherSettingsChangeSignal { get; set; }
        public bool TsOtherSettingsAutomaticRecognition { get; set; }

        public bool TsUpperMenu { get; set; }
        public bool TsSyncTime { get; set; }
        public bool TsRestart { get; set; }
        public bool TsShutdown { get; set; }
        public bool TsRemoveAdmins { get; set; }
        public bool TsFlushAllData { get; set; }
        public bool TsUpperMenuUpdateFirmware { get; set; }
        public bool TsUpperMenuTimexUsb { get; set; }
        public bool TsStatistic { get; set; }
        public bool TsStatisticLastUpdateTime { get; set; }
        public bool TsStatisticDeviceTime { get; set; }
        public int TsStatisticUsers { get; set; }
        public int TsStatisticFingerPrints { get; set; }
        public int TsStatisticFace { get; set; }
        public int TsStatisticCards { get; set; }
        public int TsStatisticCode { get; set; }
        public int TsStatisticEvents { get; set; }
        public bool TsStatisticsAdmins { get; set; }
        public bool TsStatisticSerialNumber { get; set; }
        public bool TsStatisticSoft { get; set; }

        public int FType { get; set; }
        public int FMaxUsers { get; set; }
        public int FMaxOtpechatkov { get; set; }
        public int FMaxLico { get; set; }
        public int FMaxVein { get; set; }
        public int FMaxCards { get; set; }
        public int FKod { get; set; }
        public int FMaxRecords { get; set; }
        public bool FEthernet { get; set; }
        public bool FRS232 { get; set; }
        public bool FUSBHost { get; set; }
        public bool FUSBClient { get; set; }
        public bool FWiegandOut { get; set; }
        public bool FWiegandIn { get; set; }
        public bool FWebServer { get; set; }
        public bool FVvodOtpechatkovIzTimex { get; set; }
        public int FMaxFuncButtons { get; set; }
        public int FTochekRegistracii { get; set; }
        public bool FVivodPhoto { get; set; }
        public bool FVivodImeni { get; set; }
        public bool FZahvatPhoto { get; set; }
        public bool FGolosovieSoobsheniya { get; set; }
        public bool FSMS { get; set; }
        public bool FKodRabot { get; set; }
        public bool FSignalSmeni { get; set; }
        public bool FLetoZima { get; set; }

        public bool FSkud { get; set; }
        public int FPunktovDostupa { get; set; }
        public int FProx { get; set; }
        public bool FLocalniyZPP { get; set; }
        public bool FIspSchKlav { get; set; }
        public bool FPrazdniki { get; set; }
        public int FVremennieZoni { get; set; }
        public bool FGruppiDostupa { get; set; }
        public bool FFingerPrintPrinujdenie { get; set; }
        public bool FCodPrinuzhd { get; set; }
        public bool FZamok { get; set; }
        public bool FTrevojniyVihod { get; set; }
        public bool FAvailableInFreeVersion { get; set; }
        public bool FAlgorithm9 { get; set; }
        public bool FAlgorithm10 { get; set; }
        public bool FImportTemplates { get; set; }
        public int FDopVhodov { get; set; }
        public int FDopVihodov { get; set; }
        public bool FIdentificationTypes { get; set; }
        public bool FUsingZpp { get; set; }
        public bool FDisplayNameOnTerminal { get; set; }
        public bool FAccessToTerminal { get; set; }
        public int FVeinPerEmp { get; set; }
        public bool FCustomWiegand { get; set; }

        public bool PdPunktiDostupa { get; set; }
        public bool PdSettings { get; set; }
        public bool PdSettingsDoorWorkByTimeZone { get; set; }
        public bool PdSettingsDoorUnlockByTimeZone { get; set; }
        public bool PdSettingsDoorBlockByTimeZone { get; set; }
        public bool PdSettingsDoorUnlockTimeout { get; set; }
        public bool PdSettingsReadingDelay { get; set; }
        public bool PdSettingsIdentificationMode { get; set; }
        public bool PdSettingsVedomiy { get; set; }
        public bool PdSettingsFirstCardOpenDoor { get; set; }
        public bool PdDoorMonitoring { get; set; }
        public bool PdDoorMonitoringSensorType { get; set; }
        public bool PdDoorMonitoringCloseDoorBySensor { get; set; }
        public bool PdDoorMonitoringAlarmOpenDoor { get; set; }
        public bool PdDoorMonitoringAlarmExitInterval { get; set; }
        public bool PdDopSchitivatel { get; set; }
        public bool PdDopSchitivatelVedomiy { get; set; }
        public bool PdDopSchitivatelWorkMode { get; set; }
        public bool PdDopSchitivatelName { get; set; }
        public bool PdExtra { get; set; }
        public bool PdExtraAccessCodeUnderForce { get; set; }
        public bool PdExtraManualExitOnAlarmCount { get; set; }
        public bool PdExtraEmergencyCode { get; set; }
        public bool PdExtraMultiCardOpenDoor { get; set; }
        public bool PdDoorOpen { get; set; }
        public bool PdDoorBlock { get; set; }
        public bool PdDoorUnblock { get; set; }
        public bool PdVedomiyIdentificationMode { get; set; }
        public bool PdDopSchitivatelIdentificationMode { get; set; }

        public bool RiFingerprintOrCodeOrCard { get; set; }
        public bool RiFingerprint { get; set; }
        public bool RiPin { get; set; }
        public bool RiCode { get; set; }
        public bool RiCard { get; set; }
        public bool RiFingerprintAndCard { get; set; }
        public bool RiFingerprintOrCode { get; set; }
        public bool RiFingerprintOrCard { get; set; }
        public bool RiCodeOrCard { get; set; }
        public bool RiPinAndFingerprint { get; set; }
        public bool RiFingerprintAndCode { get; set; }
        public bool RiCodeAndCard { get; set; }
        public bool RiFingerprintAndCodeAndCard { get; set; }
        public bool RiPinAndFingerprintAndCode { get; set; }
        public bool RiFingerprintAndCardOrPin { get; set; }
        public bool RiFace { get; set; }
        public bool RiFaceAndCode { get; set; }
        public bool RiFaceAndCard { get; set; }
        public bool RiFaceAndCodeAndCard { get; set; }
        public bool RiFaceOrCodeOrCard { get; set; }
        public bool RiVein { get; set; }
        public bool RiVeinAndCard { get; set; }
        public bool RiVeinAndCode { get; set; }
        public bool RiVeinAndCardAndCode { get; set; }
        public bool RiVeinOrCodeOrCard { get; set; }

        public string RiDefault { get; set; }
        public bool RiEnabled { get; set; }
        public bool RiSupportModes { get; set; }

        public bool FShluz { get; set; }

        public bool FAvailableAC { get; set; }
        public bool FAvailableTA { get; set; }
    }
}