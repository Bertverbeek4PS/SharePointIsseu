codeunit 50100 "Document Service 4PS"
{
    // Provides functions for the storage of documents to online services such as O365 (Office 365).


    trigger OnRun()
    begin
    end;

    var
        NoConfigErr: Label 'No online document configuration was found.';
        MultipleConfigsErr: Label 'More than one online document configuration was found.';
        SourceFileNotFoundErr: Label 'Cannot open the specified document from the following location: %1 due to the following error: %2.', Comment = '%1=Full path to the file on disk;%2=the detailed error describing why the document could not be accessed.';
        RequiredSourceNameErr: Label 'You must specify a source path for the document.';
        DocumentService: DotNet IDocumentService;
        DocumentServiceFactory: DotNet DocumentServiceFactory;
        ClientContext: DotNet ClientContext0;
        FileManagement: Codeunit "File Management";
        ServiceType: Text;
        LastServiceType: Text;
        RequiredTargetNameErr: Label 'You must specify a name for the document.';
        RequiredTargetURIErr: Label 'You must specify the URI that you want to open.';
        ValidateConnectionErr: Label 'Cannot connect because the user name and password have not been specified, or because the connection was canceled.';
        AccessTokenErrMsg: Label 'Failed to acquire an access token.';
        MissingClientIdOrSecretErr: Label 'The client ID or client secret have not been initialized.';
        SharePointIsoStorageSecretNotConfiguredErr: Label 'Client secret for SharePoint has not been configured.';
        SharePointIsoStorageSecretNotConfiguredLbl: Label 'Client secret for SharePoint has not been configured.', Locked = true;
        AuthTokenOrCodeNotReceivedErr: Label 'No access token or authorization error code received. The authorization failure error is: %1.', Comment = '%1=Authentiaction Failure Error', Locked = true;
        AccessTokenAcquiredFromCacheErr: Label 'The attempt to acquire the access token form cache has failed.', Locked = false;
        OAuthAuthorityUrlLbl: Label 'https://login.microsoftonline.com/common/oauth2', Locked = true;
        SharePointTelemetryCategoryTxt: Label 'AL Sharepoint Integration', Locked = true;
        SharePointClientIdAKVSecretNameLbl: Label 'sharepoint-clientid', Locked = true;
        SharePointClientSecretAKVSecretNameLbl: Label 'sharepoint-clientsecret', Locked = true;
        MissingClientIdTelemetryTxt: Label 'The client ID has not been initialized.', Locked = true;
        MissingClientSecretTelemetryTxt: Label 'The client secret has not been initialized.', Locked = true;
        InitializedClientIdTelemetryTxt: Label 'The client ID has been initialized.', Locked = true;
        InitializedClientSecretTelemetryTxt: Label 'The client secret has been initialized.', Locked = true;
        UseDocumentService4PS: Boolean;
        DoNotShowError: Boolean;

    [Scope('OnPrem')]
    procedure TestConnection()
    var
        DocumentServiceRec: Record "Document Service";
        DocumentServiceHelper: DotNet NavDocumentServiceHelper;
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
    begin
        // Tests connectivity to the Document Service using the current configuration in Dynamics NAV.
        // An error occurrs if unable to successfully connect.
        if not IsConfigured then
            Error(NoConfigErr);
        DocumentServiceHelper.Reset();
        SetDocumentService;
        SetProperties(false);

        if not UseDocumentService4PS then begin//**4PS.n
            if DocumentServiceRec.FindFirst() then
                if DocumentServiceRec."Authentication Type" = DocumentServiceRec."Authentication Type"::Legacy then
                    if IsNull(DocumentService.Credentials) then
                        Error(ValidateConnectionErr);
        end else begin
            if DocumentServiceSetup4PS.FindFirst() then
                if DocumentServiceSetup4PS."Authentication Type" = DocumentServiceSetup4PS."Authentication Type"::Legacy then
                    if IsNull(DocumentService.Credentials) then
                        Error(ValidateConnectionErr);
        end;

        DocumentService.ValidateConnection;
        CheckError;
    end;

    [Scope('OnPrem')]
    procedure SaveFile(SourcePath: Text; TargetName: Text; Overwrite: Boolean): Text
    var
        SourceFile: File;
        SourceStream: InStream;
    begin
        // Saves a file to the Document Service using the configured location specified in Dynamics NAV.
        // SourcePath: The path to a physical file on the Dynamics NAV server.
        // TargetName: The name which will be given to the file saved to the Document Service.
        // Overwrite: TRUE if the target file should be overwritten.
        // - An error is shown if Overwrite is FALSE and a file with that name already exists.
        // Returns: A URI to the file on the Document Service.

        if SourcePath = '' then
            Error(RequiredSourceNameErr);

        if TargetName = '' then
            Error(RequiredTargetNameErr);

        if not IsConfigured then
            Error(NoConfigErr);

        if not SourceFile.Open(SourcePath) then
            Error(SourceFileNotFoundErr, SourcePath, GetLastErrorText);

        SourceFile.CreateInStream(SourceStream);

        exit(SaveStream(SourceStream, TargetName, Overwrite));
    end;

    procedure IsConfigured(): Boolean
    var
        DocumentServiceRec: Record "Document Service";
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
    begin
        // Returns TRUE if Dynamics NAV has been configured with a Document Service.

        //**4PS.sn
        if UseDocumentService4PS then begin
            if not DocumentServiceSetup4PS.Get then
                exit(false);
            if (DocumentServiceSetup4PS."SharePoint Site URL" = '') or
               (DocumentServiceSetup4PS."Document Library" = '') or    //**4PS.n mvdbovenkamp
               (DocumentServiceSetup4PS.Folder = '')
            then
                exit(false);
            exit(true);
        end;
        //**4PS.en

        with DocumentServiceRec do begin
            if Count > 1 then
                Error(MultipleConfigsErr);

            if not FindFirst then
                exit(false);

            if (Location = '') or (Folder = '') then
                exit(false);
        end;

        exit(true);
    end;

    [Scope('OnPrem')]
    procedure IsServiceUri(TargetURI: Text): Boolean
    var
        DocumentServiceRec: Record "Document Service";
        IsValid: Boolean;
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
    begin
        // Returns TRUE if the TargetURI points to a location on the currently-configured Document Service.

        if TargetURI = '' then
            exit(false);

        //**4PS.sn
        if UseDocumentService4PS then begin
            if DocumentServiceSetup4PS.Get then
                if DocumentServiceSetup4PS."SharePoint Site URL" <> '' then begin
                    SetDocumentService;
                    SetProperties(true);
                    IsValid := DocumentService.IsValidUri(TargetURI);
                    if DoNotShowError then begin
                        if HasError then
                            exit(false);
                    end else
                        CheckError;
                    exit(IsValid);
                end;
            exit(false);
        end;
        //**4PS.en

        with DocumentServiceRec do begin
            if FindLast then
                if Location <> '' then begin
                    SetDocumentService;
                    SetProperties(true);
                    IsValid := DocumentService.IsValidUri(TargetURI);
                    CheckError;
                    exit(IsValid);
                end
        end;

        exit(false);
    end;

    [Scope('OnPrem')]
    procedure SetServiceType(RequestedServiceType: Text)
    var
        DocumentServiceHelper: DotNet NavDocumentServiceHelper;
    begin
        // Sets the type name of the Document Service.
        // The type must match the DocumentServiceMetadata attribute value on the IDocumentServiceHandler interface
        // exposed by at least one assembly in the Server installation folder.
        // By default, Dynamics NAV uses the SharePoint Online Document Service with type named 'SHAREPOINTONLINE'.
        ServiceType := RequestedServiceType;
        DocumentServiceHelper.SetDocumentServiceType(RequestedServiceType);
    end;

    procedure GetServiceType(): Text
    begin
        // Gets the name of the current Document Service.

        exit(ServiceType);
    end;

    [Scope('OnPrem')]
    procedure OpenDocument(TargetURI: Text)
    begin
        // Navigates to the specified URI on the Document Service from the client device.

        if TargetURI = '' then
            Error(RequiredTargetURIErr);

        if not IsConfigured then
            Error(NoConfigErr);

        SetDocumentService;
        HyperLink(DocumentService.GenerateViewableDocumentAddress(TargetURI));
        CheckError;
    end;

    [NonDebuggable]
    local procedure SetProperties(GetTokenFromCache: Boolean)
    var
        DocumentServiceRec: Record "Document Service";
        DocumentServiceHelper: DotNet NavDocumentServiceHelper;
        AccessToken: Text;
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
    begin
        //**4PS.sn
        if UseDocumentService4PS then begin
            DocumentServiceSetup4PS.Get;
            DocumentService.Properties.SetProperty(
              DocumentServiceRec.FieldName(Location), DocumentServiceSetup4PS."SharePoint Site URL");
            DocumentService.Properties.SetProperty(
              DocumentServiceRec.FieldName("User Name"), DocumentServiceSetup4PS."User Name");
            DocumentService.Properties.SetProperty(
              DocumentServiceRec.FieldName("Document Repository"), DocumentServiceSetup4PS."Document Library");
            DocumentService.Properties.SetProperty(
              DocumentServiceRec.FieldName(Folder), DocumentServiceSetup4PS.Folder);
            DocumentService.Properties.SetProperty(
              DocumentServiceRec.FieldName("Authentication Type"), DocumentServiceSetup4PS."Authentication Type");

            if (DocumentServiceSetup4PS."Authentication Type" = DocumentServiceSetup4PS."Authentication Type"::Legacy) then begin
                DocumentService.Properties.SetProperty(
                  DocumentServiceRec.FieldName(Password), DocumentServiceSetup4PS.Password);
                DocumentService.Credentials := DocumentServiceHelper.ProvideCredentials;
            end else begin
                GetAccessToken(DocumentServiceSetup4PS."SharePoint Site URL", AccessToken, GetTokenFromCache);
                DocumentService.Properties.SetProperty('Token', AccessToken);
            end;

            if not (DocumentServiceHelper.LastErrorMessage = '') then
                Error(DocumentServiceHelper.LastErrorMessage);
            exit;
        end;
        //**4PS.en

        with DocumentServiceRec do begin
            if not FindFirst then
                Error(NoConfigErr);

            // The Document Service will throw an exception if the property is not known to the service type provider.
            DocumentService.Properties.SetProperty(FieldName(Description), Description);
            DocumentService.Properties.SetProperty(FieldName(Location), Location);
            DocumentService.Properties.SetProperty(FieldName("Document Repository"), "Document Repository");
            DocumentService.Properties.SetProperty(FieldName(Folder), Folder);
            DocumentService.Properties.SetProperty(FieldName("Authentication Type"), "Authentication Type");
            DocumentService.Properties.SetProperty(FieldName("User Name"), "User Name");

            if ("Authentication Type" = "Authentication Type"::Legacy) then begin
                DocumentService.Properties.SetProperty(FieldName(Password), Password);
                DocumentService.Credentials := DocumentServiceHelper.ProvideCredentials;
            end else begin
                GetAccessToken(Location, AccessToken, GetTokenFromCache);
                DocumentService.Properties.SetProperty('Token', AccessToken);
            end;

            if not (DocumentServiceHelper.LastErrorMessage = '') then
                Error(DocumentServiceHelper.LastErrorMessage);
        end;
    end;

    local procedure SetDocumentService()
    var
        RequestedServiceType: Text;
    begin
        // Sets the Document Service for the current Service Type, reusing an existing service if possible.

        RequestedServiceType := GetServiceType;

        if RequestedServiceType = '' then
            RequestedServiceType := 'SHAREPOINTONLINE';

        if LastServiceType <> RequestedServiceType then begin
            DocumentService := DocumentServiceFactory.CreateService(RequestedServiceType);
            LastServiceType := RequestedServiceType;
        end;
    end;

    local procedure CheckError()
    begin
        // Checks whether the Document Service received an error and displays that error to the user.

        if not IsNull(DocumentService.LastError) and (DocumentService.LastError.Message <> '') then
            Error(DocumentService.LastError.Message);
    end;

    [Scope('OnPrem')]
    procedure SaveStream(Stream: InStream; TargetName: Text; Overwrite: Boolean): Text
    var
        DocumentURI: Text;
        FileManagement: Codeunit "File Management";
        SubfolderName: Text;
    begin
        // Saves a stream to the Document Service using the configured location specified in Dynamics NAV.
        SetDocumentService;
        SetProperties(true);

        //**4PS.sn
        if UseDocumentService4PS then begin
            GetSubfolderName(TargetName, SubfolderName);
            SetSubfolder(SubfolderName);
            TargetName := FileManagement.GetFileName(SubfolderName);
        end;
        //**4PS.en
        DocumentURI := DocumentService.Save(Stream, TargetName, Overwrite);
        CheckError;

        exit(DocumentURI);
    end;

    [NonDebuggable]
    local procedure GetAccessToken(Location: Text; var AccessToken: Text; GetTokenFromCache: Boolean)
    var
        OAuth2: Codeunit OAuth2;
        EnvironmentInformation: Codeunit "Environment Information";
        PromptInteraction: Enum "Prompt Interaction";
        ClientId: Text;
        ClientSecret: Text;
        RedirectURL: Text;
        ResourceURL: Text;
        AuthError: Text;
    begin
        ResourceURL := GetResourceUrl(Location);

        if EnvironmentInformation.IsSaaSInfrastructure() then begin
            OAuth2.AcquireOnBehalfOfToken('', ResourceURL, AccessToken);
            exit;
        end;

        ClientId := GetClientId();
        ClientSecret := GetClientSecret();
        RedirectURL := GetRedirectURL();

        if GetTokenFromCache then
            OAuth2.AcquireAuthorizationCodeTokenFromCache(ClientId, ClientSecret, RedirectURL, OAuthAuthorityUrlLbl, ResourceURL, AccessToken);

        if AccessToken <> '' then
            exit;

        Session.LogMessage('0000DB7', AccessTokenAcquiredFromCacheErr, Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
        OAuth2.AcquireTokenByAuthorizationCode(
                    ClientId,
                    ClientSecret,
                    OAuthAuthorityUrlLbl,
                    RedirectURL,
                    ResourceURL,
                    PromptInteraction::Consent,
                    AccessToken,
                    AuthError
                );

        if AccessToken = '' then begin
            Session.LogMessage('0000DB8', StrSubstNo(AuthTokenOrCodeNotReceivedErr, AuthError), Verbosity::Error, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
            Error(AccessTokenErrMsg);
        end;
    end;

    [NonDebuggable]
    local procedure GetClientId(): Text
    var
        DocumentServiceRec: Record "Document Service";
        AzureKeyVault: Codeunit "Azure Key Vault";
        EnvironmentInformation: Codeunit "Environment Information";
        ClientId: Text;
    begin
        if EnvironmentInformation.IsSaaSInfrastructure() then
            if not AzureKeyVault.GetAzureKeyVaultSecret(SharePointClientIdAKVSecretNameLbl, ClientId) then
                Session.LogMessage('0000DB9', MissingClientIdTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt)
            else begin
                Session.LogMessage('0000DBA', InitializedClientIdTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
                exit(ClientId);
            end;

        if DocumentServiceRec.FindFirst() then begin
            ClientId := DocumentServiceRec."Client Id";
            OnGetSharePointClientId(ClientId);
            if ClientId <> '' then begin
                Session.LogMessage('0000DBB', InitializedClientIdTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
                exit(ClientId);
            end;
        end;

        Error(MissingClientIdOrSecretErr);
    end;

    [NonDebuggable]
    local procedure GetClientSecret(): Text
    var
        DocumentServiceRec: Record "Document Service";
        AzureKeyVault: Codeunit "Azure Key Vault";
        EnvironmentInformation: Codeunit "Environment Information";
        ClientSecret: Text;
    begin
        if EnvironmentInformation.IsSaaSInfrastructure() then
            if not AzureKeyVault.GetAzureKeyVaultSecret(SharePointClientSecretAKVSecretNameLbl, ClientSecret) then
                Session.LogMessage('0000DBC', MissingClientSecretTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt)
            else begin
                Session.LogMessage('0000DBD', InitializedClientSecretTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
                exit(ClientSecret);
            end;

        if not DocumentServiceRec.IsEmpty() then begin
            ClientSecret := GetClientSecretFromIsolatedStorage();
            if ClientSecret <> '' then begin
                Session.LogMessage('0000DBE', InitializedClientSecretTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
                exit(ClientSecret);
            end;
        end;

        OnGetSharePointClientSecret(ClientSecret);
        if ClientSecret <> '' then begin
            Session.LogMessage('0000DBF', InitializedClientSecretTelemetryTxt, Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
            exit(ClientSecret);
        end;

        Error(MissingClientIdOrSecretErr);
    end;

    [Scope('OnPrem')]
    [NonDebuggable]
    local procedure GetRedirectURL(): Text
    var
        DocumentServiceRec: Record "Document Service";
        EnvironmentInformation: Codeunit "Environment Information";
        RedirectURL: Text;
    begin
        if EnvironmentInformation.IsSaaSInfrastructure() then
            exit(RedirectURL);

        if DocumentServiceRec.FindFirst() then
            RedirectURL := DocumentServiceRec."Redirect URL";

        if RedirectURL = '' then
            OnGetSharePointRedirectURL(RedirectURL);

        exit(RedirectURL);
    end;

    [NonDebuggable]
    local procedure GetResourceUrl(Location: Text): Text
    begin
        exit(Location.Substring(1, Location.IndexOf('.com') + 3));
    end;


    [Scope('OnPrem')]
    [NonDebuggable]
    internal procedure SetClientSecret(ClientSecret: Text)
    var
        DocumentServiceRec: Record "Document Service";
        IsolatedStorageManagement: Codeunit "Isolated Storage Management";
    begin
        if not DocumentServiceRec.FindFirst() then
            Error(NoConfigErr);

        if ClientSecret = '' then
            if not IsNullGuid(DocumentServiceRec."Client Secret Key") then begin
                IsolatedStorageManagement.Delete(DocumentServiceRec."Client Secret Key", DATASCOPE::Company);
                exit;
            end;

        if IsNullGuid(DocumentServiceRec."Client Secret Key") then begin
            DocumentServiceRec."Client Secret Key" := CreateGuid();
            DocumentServiceRec.Modify();
        end;

        IsolatedStorageManagement.Set(DocumentServiceRec."Client Secret Key", ClientSecret, DATASCOPE::Company);
    end;

    [Scope('OnPrem')]
    [NonDebuggable]
    local procedure GetClientSecretFromIsolatedStorage(): Text
    var
        DocumentServiceRec: Record "Document Service";
        IsolatedStorageManagement: Codeunit "Isolated Storage Management";
        ClientSecret: Text;
    begin
        if not DocumentServiceRec.FindFirst() then
            Error(NoConfigErr);

        if IsNullGuid(DocumentServiceRec."Client Secret Key") or
           not IsolatedStorage.Contains(DocumentServiceRec."Client Secret Key", DATASCOPE::Company)
        then begin
            Session.LogMessage('0000DBG', SharePointIsoStorageSecretNotConfiguredLbl, Verbosity::Warning, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', SharePointTelemetryCategoryTxt);
            Error(SharePointIsoStorageSecretNotConfiguredErr);
        end;

        IsolatedStorageManagement.Get(DocumentServiceRec."Client Secret Key", DATASCOPE::Company, ClientSecret);
        exit(ClientSecret);
    end;

    [Scope('OnPrem')]
    [NonDebuggable]
    [TryFunction]
    internal procedure TryGetClientSecretFromIsolatedStorage(var ClientSecret: Text)
    begin
        ClientSecret := GetClientSecretFromIsolatedStorage();
    end;

    [EventSubscriber(ObjectType::Codeunit, 2000000006, 'OnOpenInExcel', '', false, false)]
    [NonDebuggable]
    local procedure OnTryAcquireAccessTokenOnOpenInExcel(Location: Text)
    var
        Token: Text;
    begin
        GetAccessToken(Location, Token, true);
        Session.SetDocumentServiceToken(Token);
    end;

    [IntegrationEvent(false, false)]
    local procedure OnGetSharePointClientId(var ClientId: Text)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnGetSharePointClientSecret(var ClientSecret: Text)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnGetSharePointRedirectURL(var RedirectURL: Text)
    begin
    end;

    procedure SaveStreamByOutputStream(MemoryStream: DotNet MemoryStream; TargetName: Text; Overwrite: Boolean): Text
    var
        DocumentURI: Text;
        FileManagement: Codeunit "File Management";
        OStream: OutStream;
        DocumentServiceHelper: DotNet NavDocumentServiceHelper;
    begin
        //**4PS
        DocumentServiceHelper.Reset();
        SetDocumentService;
        SetProperties(true);

        if UseDocumentService4PS then begin
            SetSubfolder(TargetName);
            TargetName := FileManagement.GetFileName(TargetName);
        end;

        OStream := MemoryStream;
        DocumentURI := DocumentService.Save(OStream, TargetName, Overwrite);
        CheckError;

        exit(DocumentURI);
    end;

    procedure SaveStreamByOutputStream(OStream: OutStream; TargetName: Text; Overwrite: Boolean): Text
    var
        DocumentURI: Text;
        FileManagement: Codeunit "File Management";
    begin
        //**4PS
        SetDocumentService;
        SetProperties(true);

        if UseDocumentService4PS then begin
            SetSubfolder(TargetName);
            TargetName := FileManagement.GetFileName(TargetName);
        end;

        DocumentURI := DocumentService.Save(OStream, TargetName, Overwrite);
        CheckError;

        exit(DocumentURI);
    end;

    local procedure SetSubfolder(TargetName: Text)
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        DocumentService2: Record "Document Service";
        FileManagement: Codeunit "File Management";
        Subfolder: Text;
    begin
        //**4PS
        DocumentServiceSetup4PS.Get;
        Subfolder := FileManagement.GetDirectoryName(TargetName);
        if (Subfolder <> '') then begin
            if (CopyStr(Subfolder, 1, 1) = '\') then
                Subfolder := CopyStr(Subfolder, 2);
            DocumentService.Properties.SetProperty(
              DocumentService2.FieldName(Folder), DocumentServiceSetup4PS.Folder + '\' + Subfolder);
        end;
    end;

    procedure DownloadFileFromCloudToServerSilent(DocumentURI: Text) ServerFileName: Text
    var
        WebClient: DotNet WebClient;
        SecureString: DotNet SecureString;
        Uri: DotNet Uri;
        SharePointOnlineCredentials: DotNet SharePointOnlineCredentials;
        HttpRequestHeader: DotNet HttpRequestHeader;
        FileHelper: DotNet File;
        DateTimeHelper: DotNet DateTime;
        DocumentServiceRec: Record "Document Service";
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        FileManagement: Codeunit "File Management";
        LastModifiedDateTime: DateTime;
        PassWord: Text;
        UserName: Text;
        LastModifiedText: Text;
        I: Integer;
    begin
        //**4PS
        if DocumentURI = '' then
            Error(RequiredTargetURIErr);

        if not IsConfigured then
            Error(NoConfigErr);

        if UseDocumentService4PS then begin
            DocumentServiceSetup4PS.Get;
            PassWord := DocumentServiceSetup4PS.Password;
            UserName := DocumentServiceSetup4PS."User Name";
        end else begin
            DocumentServiceRec.FindFirst;
            PassWord := DocumentServiceRec.Password;
            UserName := DocumentServiceRec."User Name";
        end;

        SecureString := SecureString.SecureString;
        while (I < StrLen(PassWord)) do begin
            SecureString.AppendChar(PassWord[I + 1]);
            I += 1;
        end;

        if IsRelativePath(DocumentURI) then
            ConvertRelativePathToURI(DocumentURI, DocumentURI);

        Uri := Uri.Uri(DocumentURI);
        SharePointOnlineCredentials := SharePointOnlineCredentials.SharePointOnlineCredentials(
          UserName, SecureString);
        WebClient := WebClient.WebClient;
        WebClient.Headers.Add(
          HttpRequestHeader.Cookie, SharePointOnlineCredentials.GetAuthenticationCookie(Uri));
        ServerFileName := FileManagement.ServerTempFileName(FileManagement.GetExtension(DocumentURI));
        WebClient.DownloadFile(DocumentURI, ServerFileName);
        //Last modified date of server file is set, to make sure that
        //DocumentProperties.GetPropertiesFromFile returns the correct ModifyDateTime.
        LastModifiedText := WebClient.ResponseHeaders.Get('Last-Modified');
        if DateTimeHelper.TryParse(LastModifiedText, LastModifiedDateTime) then
            FileHelper.SetLastWriteTime(ServerFileName, LastModifiedDateTime);
    end;

    procedure DownloadFileFromCloudToClientSilent(DocumentURI: Text) ClientFileName: Text
    var
        FileManagement: Codeunit "File Management";
        TempServerFileName: Text;
    begin
        //**4PS
        TempServerFileName := DownloadFileFromCloudToServerSilent(DocumentURI);
        ClientFileName := FileManagement.DownloadTempFile(TempServerFileName);
        Erase(TempServerFileName);
    end;

    procedure DownloadFileFromCloudToTempBlobByRelativePath(RelativePath: Text; var TempBlob: Codeunit "Temp Blob")
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        URI: DotNet Uri;
        FileManagement: Codeunit "File Management";
        TempServerFileName: Text;
        DocumentURI: Text;
    begin
        //**4PS
        DocumentServiceSetup4PS.Get();
        DocumentServiceSetup4PS.TestField("SharePoint Site URL");
        URI := URI.Uri(DocumentServiceSetup4PS."SharePoint Site URL");

        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        DocumentURI := 'https://' + URI.Host + RelativePath;

        SetUseDocumentService4PS(true);

        TempServerFileName := DownloadFileFromCloudToServerSilent(DocumentURI);
        FileManagement.BLOBImportFromServerFile(TempBlob, TempServerFileName);
        Erase(TempServerFileName);
    end;

    procedure DownloadFileFromCloudStreamed(DocumentURI: Text; var Base64: BigText): Boolean
    var
        WebClient: DotNet WebClient;
        MemoryStream: DotNet MemoryStream;
        Stream: DotNet Stream;
        Bytes: DotNet Array;
        Convert: DotNet Convert;
    begin
        //**4PS
        WebClient := WebClient.WebClient;
        SetWebClientCredentials(WebClient, DocumentURI);
        Stream := WebClient.OpenRead(DocumentURI);
        MemoryStream := MemoryStream.MemoryStream;
        Stream.CopyTo(MemoryStream);
        Bytes := MemoryStream.ToArray;
        Base64.AddText(Convert.ToBase64String(Bytes));
        exit(true);
    end;

    procedure DownloadFileFromCloudToMemoryStream(DocumentURI: Text; var MemoryStream: DotNet MemoryStream)
    var
        WebClient: DotNet WebClient;
        Stream: DotNet Stream;
    begin
        //**4PS
        WebClient := WebClient.WebClient;
        SetWebClientCredentials(WebClient, DocumentURI);
        Stream := WebClient.OpenRead(DocumentURI);
        Stream.CopyTo(MemoryStream);
    end;

    procedure UploadFileFromServerToCloudSilent(ServerFileName: Text; TargetFolder: Text; ShortFileName: Text) DocumentURI: Text
    var
    //FileStorageManagement: Codeunit "File Storage Management";
    begin
        //**4PS
        //FileStorageManagement.AddSlashToStorageDirectory(2, TargetFolder);
        //TestIsConfigured;
        //DocumentURI := SaveFile(ServerFileName, TargetFolder + ShortFileName, true);
    end;

    procedure UploadFileFromClientToCloudSilent(ClientFileName: Text; TargetFolder: Text) DocumentURI: Text
    var
        ShortFileName: Text;
        ServerFileName: Text;
    begin
        //**4PS
        ShortFileName := FileManagement.GetFileName(ClientFileName);
        ServerFileName := FileManagement.UploadFileSilent(ClientFileName);
        DocumentURI := UploadFileFromServerToCloudSilent(ServerFileName, TargetFolder, ShortFileName);
        Erase(ServerFileName);
    end;

    procedure UploadFileFromClientToCloud(TargetFolder: Text; Extension: Text) DocumentURI: Text
    var
        ServerFileName: Text;
        ShortFileName: Text;
    begin
        //**4PS
        if Extension <> '' then
            ServerFileName := FileManagement.UploadFile('', '*.' + Extension)
        else
            ServerFileName := FileManagement.UploadFile('', '');
        if ServerFileName = '' then
            exit;
        ShortFileName := FileManagement.GetFileName(ServerFileName);
        DocumentURI := UploadFileFromServerToCloudSilent(ServerFileName, TargetFolder, ShortFileName);
        Erase(ServerFileName);
    end;

    procedure TestIsConfigured()
    begin
        //**4PS
        if not IsConfigured then
            Error(NoConfigErr);
    end;

    procedure SetUseDocumentService4PS(UseDocumentService4PS2: Boolean)
    begin
        //**4PS
        UseDocumentService4PS := UseDocumentService4PS2;
    end;

    procedure SetDoNotShowError(DoNotShowErrorNew: Boolean)
    begin
        //**4PS
        DoNotShowError := DoNotShowErrorNew;
    end;

    local procedure HasError(): Boolean
    begin
        //**4PS
        if not IsNull(DocumentService.LastError) and (DocumentService.LastError.Message <> '') then
            exit(true);

        exit(false);
    end;

    procedure CanConnect(): Boolean
    var
        DocumentServiceHelper: DotNet NavDocumentServiceHelper;
    begin
        //**4PS
        if not IsConfigured then
            exit(false);
        DocumentServiceHelper.Reset;
        SetDocumentService;
        SetProperties(false);

        if not UseDocumentService4PS then
            if IsNull(DocumentService.Credentials) then
                exit(false);

        DocumentService.ValidateConnection();
        exit(not HasError);
    end;

    procedure GetDocumentServiceURL() DocumentServiceURL: Text
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
    begin
        //**4PS
        DocumentServiceURL := '';
        if not UseDocumentService4PS then
            exit;
        if not CanConnect then
            exit;
        DocumentServiceSetup4PS.Get;
        DocumentServiceURL := DocumentServiceSetup4PS."SharePoint Site URL";
        if (CopyStr(DocumentServiceSetup4PS."SharePoint Site URL", StrLen(DocumentServiceSetup4PS."SharePoint Site URL")) <> '/') and
           (CopyStr(DocumentServiceSetup4PS."Document Library", 1, 1) <> '/')
        then
            DocumentServiceURL := DocumentServiceURL + '/';
        DocumentServiceURL := DocumentServiceURL + DocumentServiceSetup4PS."Document Library";
        if DocumentServiceSetup4PS.Folder = '' then
            exit;
        if (CopyStr(DocumentServiceSetup4PS."Document Library", StrLen(DocumentServiceSetup4PS."Document Library")) <> '/') and
           (CopyStr(DocumentServiceSetup4PS.Folder, 1, 1) <> '/')
        then
            DocumentServiceURL := DocumentServiceURL + '/';
        DocumentServiceURL := DocumentServiceURL + DocumentServiceSetup4PS.Folder;
    end;

    local procedure SetSecureString(var SecureString: DotNet SecureString; PassWord: Text)
    var
        I: Integer;
    begin
        //**4PS
        while (I < StrLen(PassWord)) do begin
            SecureString.AppendChar(PassWord[I + 1]);
            I += 1;
        end;
    end;

    local procedure SetWebClientCredentials(var WebClient: DotNet WebClient; DocumentURI: Text)
    var
        Uri: DotNet Uri;
        SharePointOnlineCredentials: DotNet SharePointOnlineCredentials;
        HttpRequestHeader: DotNet HttpRequestHeader;
        SecureString: DotNet SecureString;
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        PassWord: Text;
        UserName: Text;
        DocumentServiceRec: Record "Document Service";
    begin
        //**4PS
        if not IsConfigured then
            Error(NoConfigErr);

        if UseDocumentService4PS then begin
            DocumentServiceSetup4PS.Get;
            PassWord := DocumentServiceSetup4PS.Password;
            UserName := DocumentServiceSetup4PS."User Name";
        end else begin
            DocumentServiceRec.FindFirst;
            PassWord := DocumentServiceRec.Password;
            UserName := DocumentServiceRec."User Name";
        end;

        Uri := Uri.Uri(DocumentURI);
        SecureString := SecureString.SecureString;
        SetSecureString(SecureString, PassWord);
        SharePointOnlineCredentials := SharePointOnlineCredentials.SharePointOnlineCredentials(UserName, SecureString);
        WebClient.Headers.Add(HttpRequestHeader.Cookie, SharePointOnlineCredentials.GetAuthenticationCookie(Uri));
    end;

    procedure UploadDocumentStreamed(var Content: InStream; TargetName: Text; Overwrite: Boolean): Text
    begin
        //**4PS
        if not IsConfigured then
            Error(NoConfigErr);

        exit(SaveStream(Content, TargetName, Overwrite));
    end;


    local procedure InitClientContext()
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        SharePointOnlineCredentials: DotNet SharePointOnlineCredentials;
        [SuppressDispose]
        SecureString: DotNet SecureString;
        PassWord: Text;
        UserName: Text;
        Site: Text;
    begin
        //**4PS
        DocumentServiceSetup4PS.Get();
        PassWord := DocumentServiceSetup4PS.Password;
        UserName := DocumentServiceSetup4PS."User Name";
        Site := DocumentServiceSetup4PS."SharePoint Site URL";

        SecureString := SecureString.SecureString;
        SetSecureString(SecureString, PassWord);

        SharePointOnlineCredentials := SharePointOnlineCredentials.SharePointOnlineCredentials(UserName, SecureString);
        ClientContext := ClientContext.ClientContext(Site);

        ClientContext.Credentials := SharePointOnlineCredentials.SharePointOnlineCredentials(UserName, SecureString);
        ClientContext.Web.Context.Credentials := SharePointOnlineCredentials.SharePointOnlineCredentials(UserName, SecureString);
    end;

    local procedure GetSubfolderName(Path: Text; var SubFolder: Text)
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        BaseUrl: Text;
    begin
        //**4PS
        DocumentServiceSetup4PS.Get();
        ConvertRelativePathToURI(Path, Path);
        BaseUrl := StrSubstNo('%1/%2/%3', DocumentServiceSetup4PS."SharePoint Site URL", DocumentServiceSetup4PS."Document Library", DocumentServiceSetup4PS.Folder);
        SubFolder := Path.Replace(BaseUrl, '');
    end;

    procedure ConvertRelativePathToURI(RelativePath: Text; var SharePointEntityURI: Text): Boolean
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        URI: DotNet Uri;
        Site: Text;
        Pos: Integer;
        Host: Text;
    begin
        //**4PS
        if RelativePath = '' then
            exit(false);
        DocumentServiceSetup4PS.Get();
        Site := DocumentServiceSetup4PS."SharePoint Site URL";

        if StrPos(RelativePath, Site) > 0 then begin
            SharePointEntityURI := RelativePath;
            exit(true);
        end;

        URI := URI.Uri(Site);
        Host := URI.Host;
        Host := Host.Replace('\', '');

        //FileManagement.AddBackSlashToDirectoryName(Host);

        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        SharePointEntityURI := 'https://' + URI.Host + RelativePath;
        exit(true);
    end;

    local procedure ConvertURIToRelativePath(SharePointEntityURI: Text; var RelativePath: Text): Boolean
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        URI: DotNet Uri;
        Site: Text;
        Pos: Integer;
    begin
        //**4PS
        if SharePointEntityURI = '' then
            exit(false);
        DocumentServiceSetup4PS.Get();
        Site := DocumentServiceSetup4PS."SharePoint Site URL";

        if StrPos(SharePointEntityURI, Site) <= 0 then
            exit(false); //False if URI is not complete.

        URI := URI.Uri(SharePointEntityURI);
        Pos := StrPos(SharePointEntityURI, URI.Host);
        Pos += StrLen(URI.Host);
        RelativePath := CopyStr(SharePointEntityURI, Pos);
        exit(true);
    end;

    procedure DeleteFile(DocumentURI: Text): Boolean
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(DocumentURI, RelativePath) then
            exit(false); //File can only be deleted when URI is complete.

        exit(DeleteFileByRelativePath(RelativePath));
    end;

    procedure DeleteFileByRelativePath(RelativePath: Text): Boolean
    var
        Web: DotNet Web0;
        File: DotNet File1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;

        File := Web.GetFileByServerRelativeUrl(RelativePath);
        File.Recycle();

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();
        exit(true);
    end;

    procedure DeleteFolderByRelativePath(RelativePath: Text): Boolean
    var
        Web: DotNet Web0;
        Folder: DotNet Folder1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;

        Folder := Web.GetFolderByServerRelativeUrl(RelativePath);
        Folder.Recycle();

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();
        exit(true);
    end;

    local procedure LoadContext(Entity: DotNet Object; EntityType: DotNet Type)
    var
        MethodInfo: DotNet MethodInfo;
        ParamsArray: DotNet Array;
        TypesArray: DotNet Array;
        NullObj: DotNet Object;
    begin
        //**4PS
        MethodInfo := ClientContext.GetType().GetMethod('Load');

        TypesArray := TypesArray.CreateInstance(GetDotNetType(EntityType), 1);
        TypesArray.SetValue(EntityType, 0);

        MethodInfo := MethodInfo.MakeGenericMethod(TypesArray);

        ParamsArray := ParamsArray.CreateInstance(GetDotNetType(Entity), 2);
        ParamsArray.SetValue(Entity, 0);
        ParamsArray.SetValue(NullObj, 1);

        MethodInfo.Invoke(ClientContext, ParamsArray);
    end;

    procedure GetFilesFromFolder(FolderURI: Text; var Files: List of [Text]): Boolean
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(FolderURI, RelativePath) then
            exit(false); //File can only be deleted when URI is complete.

        exit(GetFilesFromFolderByRelativePath(RelativePath, Files));
    end;

    procedure GetFilesFromFolderByRelativePath(RelativePath: Text; var Files: List of [Text]): Boolean
    var
        Web: DotNet Web0;
        File: DotNet File1;
        Folder: DotNet Folder1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;
        Folder := Web.GetFolderByServerRelativeUrl(RelativePath);

        LoadContext(Folder, Folder.GetType());
        LoadContext(Folder.Files, Folder.Files.GetType());

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();

        Clear(Files);

        foreach File in Folder.Files() do
            Files.Add(FileManagement.CombinePath(RelativePath, File.Name));

        exit(true);
    end;

    procedure GetSubFoldersFromFolder(FolderURI: Text; var Folders: List of [Text]): Boolean
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(FolderURI, RelativePath) then
            exit(false); //File can only be deleted when URI is complete.

        exit(GetSubFoldersFromFolderByRelativePath(RelativePath, Folders));
    end;

    procedure GetSubFoldersFromFolderByRelativePath(RelativePath: Text; var Folders: List of [Text]): Boolean
    var
        Web: DotNet Web0;
        SubFolder: DotNet Folder1;
        Folder: DotNet Folder1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;
        Folder := Web.GetFolderByServerRelativeUrl(RelativePath);

        LoadContext(Folder, Folder.GetType());
        LoadContext(Folder.Folders, Folder.Folders.GetType());

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();

        Clear(Folders);

        foreach SubFolder in Folder.Folders() do
            Folders.Add(SubFolder.Name);

        exit(true);
    end;


    procedure GetFileLastMofifiedDate(DocumentURI: Text; LastModifiedDateTime: DateTime): Boolean
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(DocumentURI, RelativePath) then
            exit(false); //File can only be deleted when URI is complete.

        exit(GetFileLastMofifiedDateByRelativePath(DocumentURI, LastModifiedDateTime));
    end;

    procedure GetFileLastMofifiedDateByRelativePath(RelativePath: Text; var LastModifiedDateTime: DateTime): Boolean
    var
        Web: DotNet Web0;
        File: DotNet File1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;

        File := Web.GetFileByServerRelativeUrl(RelativePath);
        LoadContext(File, File.GetType());
        ClientContext.ExecuteQuery();
        LastModifiedDateTime := File.TimeLastModified;
        ClientContext.Dispose;
        exit(true);
    end;

    procedure MoveFile(DocumentURI: Text; DestinationURI: Text): Boolean
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(DocumentURI, RelativePath) then
            exit(false); //File can only be deleted when URI is complete.

        exit(MoveFileByRelativePath(RelativePath, DestinationURI));
    end;

    procedure MoveFileByRelativePath(RelativePath: Text; DestinationURI: Text): Boolean
    var
        Web: DotNet Web0;
        File: DotNet File1;
        MoveOperations: DotNet MoveOperations;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        if not DestinationURI.StartsWith('/') then
            DestinationURI := '/' + DestinationURI;

        Web := ClientContext.Web;

        File := Web.GetFileByServerRelativeUrl(RelativePath);
        LoadContext(File, File.GetType());
        ClientContext.ExecuteQuery();
        File.MoveTo(FileManagement.CombinePath(DestinationURI, File.Name), MoveOperations::Overwrite);

        ClientContext.ExecuteQuery();
        ClientContext.Dispose;
        exit(true);
    end;

    procedure CreateFolderByRelativePath(RelativePath: Text): Boolean
    var
        FileManagement: Codeunit "File Management";
        Web: DotNet Web0;
        Folder: DotNet Folder1;
        FolderAbsPathName: Text;
        FolderPath: Text;
    begin
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;

        FolderAbsPathName := DelChr(RelativePath, '>', '/');
        FolderPath := FileManagement.GetDirectoryName(FolderAbsPathName);

        Folder := Web.GetFolderByServerRelativeUrl(FolderPath);
        Folder.Folders.Add(FileManagement.GetFileName(FolderAbsPathName));
        Folder.Context.ExecuteQuery();

        ClientContext.Dispose();
        exit(true);
    end;

    procedure CreateFile(var IStream: InStream; DocumentURI: Text; Overwrite: Boolean)
    var
        RelativePath: Text;
    begin
        //**4PS
        if not ConvertURIToRelativePath(DocumentURI, RelativePath) then
            exit; //File can only be deleted when URI is complete.

        CreateFileByRelativePath(IStream, RelativePath, Overwrite);
    end;

    procedure CreateFileByRelativePath(var IStream: InStream; var RelativePath: Text; Overwrite: Boolean)
    var
        Web: DotNet Web0;
        Folder: DotNet Folder1;
        File: DotNet File1;
        FileManagement: Codeunit "File Management";
        FileCreationInformation: DotNet FileCreationInformation;
        MemoryStream: DotNet MemoryStream;
        FileName: Text;
        FolderName: Text;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        FileName := FileManagement.GetFileName(RelativePath);
        FolderName := FileManagement.GetDirectoryName(RelativePath);

        Web := ClientContext.Web;
        Folder := Web.GetFolderByServerRelativeUrl(FolderName);

        FileCreationInformation := FileCreationInformation.FileCreationInformation();
        FileCreationInformation.Overwrite := Overwrite;
        // CopyStream(MemoryStream, IStream);
        FileCreationInformation.ContentStream := IStream;
        FileCreationInformation.Url := FileManagement.GetFileName(RelativePath);

        File := Folder.Files.Add(FileCreationInformation);

        LoadContext(Folder, Folder.GetType());
        LoadContext(Folder.Files, Folder.Files.GetType());
        LoadContext(File, File.GetType());
        ClientContext.ExecuteQuery();
        ClientContext.Dispose();
    end;

    [TryFunction]
    procedure TryCheckDirectoryExistsByRelativePath(RelativePath: Text)
    var
        Web: DotNet Web0;
        Folder: DotNet Folder1;
        FileManagement: Codeunit "File Management";
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;
        Folder := Web.GetFolderByServerRelativeUrl(RelativePath);

        LoadContext(Folder, Folder.GetType());

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();
    end;

    [TryFunction]
    procedure TryCheckFileExistsByRelativePath(RelativePath: Text)
    var
        Web: DotNet Web0;
        File: DotNet File1;
    begin
        //**4PS
        InitClientContext();
        if not RelativePath.StartsWith('/') and IsRelativePath(RelativePath) then
            RelativePath := '/' + RelativePath;

        Web := ClientContext.Web;
        File := Web.GetFileByServerRelativeUrl(RelativePath);

        LoadContext(File, File.GetType());

        ClientContext.ExecuteQuery();
        ClientContext.Dispose();
    end;

    procedure CheckDirectoryExistsByRelativePath(RelativePath: Text) Result: Boolean
    begin
        //**4PS
        if TryCheckDirectoryExistsByRelativePath(RelativePath) then
            exit(true);
        exit(false);
    end;

    procedure CheckFileExistsByRelativePath(RelativePath: Text): Boolean
    begin
        //**4PS
        if TryCheckFileExistsByRelativePath(RelativePath) then
            exit(true);
        exit(false);
    end;

    procedure CheckEntityExistsByRelativePath(RelativePath: Text): Boolean
    begin
        //**4PS
        if TryCheckFileExistsByRelativePath(RelativePath) then
            exit(true);

        if TryCheckDirectoryExistsByRelativePath(RelativePath) then
            exit(true);

        exit(false);
    end;

    procedure GetDefaultPath(): Text
    var
        DocumentServiceSetup4PS: Record "Document Service Setup 4PS";
        URI: DotNet Uri;
    begin
        //**4PS
        DocumentServiceSetup4PS.Get();
        DocumentServiceSetup4PS.TestField("SharePoint Site URL");
        URI := URI.Uri(DocumentServiceSetup4PS."SharePoint Site URL");
        exit(URI.PathAndQuery)
    end;

    procedure IsRelativePath(RelativePath: Text): Boolean
    begin
        //**4PS
        exit(StrPos(UpperCase(RelativePath), 'HTTPS://') <> 1);
    end;
}

