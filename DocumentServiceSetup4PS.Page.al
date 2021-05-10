page 50100 "Document Service Setup 4PS"
{
    Caption = 'Document Service Setup 4PS';
    PageType = Card;
    Permissions = TableData 2000000114 = rimd;
    SourceTable = "Document Service Setup 4PS";
    UsageCategory = Administration;
    ApplicationArea = Basic, Suite;

    layout
    {
        area(content)
        {
            group(Algemeen)
            {
                Caption = 'General';
                field("SharePoint Site URL"; "SharePoint Site URL")
                {
                    ApplicationArea = Basic, Suite;
                }
                field("Document Library"; "Document Library")
                {
                    ApplicationArea = Basic, Suite;
                }
                field(Folder; Folder)
                {
                    ApplicationArea = Basic, Suite;
                }
                field("User Name"; "User Name")
                {
                    ApplicationArea = Basic, Suite;
                }
                field(Password; Password)
                {
                    ApplicationArea = Basic, Suite;
                }
                field("Default Ext. Doc. Subdir."; "Default Ext. Doc. Subdir.")
                {
                    ApplicationArea = Basic, Suite;
                }
                field("Subdir. Document Parts"; "Subdir. Document Parts")
                {
                    ApplicationArea = Basic, Suite;
                }
                field("Authentication Type"; "Authentication Type")
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Authentication Type';
                    ToolTip = 'Specifies the authentication type that will be used to connect to the SharePoint environment', Comment = 'SharePong is the name of a Microsoft service and should not be translated.';
                    trigger OnValidate()
                    var
                        User: Record User;
                    begin
                        if "Authentication Type" = "Authentication Type"::Legacy then
                            IsLegacyAuthentication := true
                        else begin
                            IsLegacyAuthentication := false;
                            if User.Get(UserSecurityId()) then
                                if User."Authentication Email" <> "User Name" then begin
                                    "User Name" := '';
                                    Modify(false);
                                    CurrPage.Update(false);
                                end;
                        end;
                    end;
                }
            }
            group("Authentication")
            {
                Caption = 'Authentication';
                Visible = not SoftwareAsAService and not IsLegacyAuthentication;
                field("Client Id"; "Client Id")
                {
                    ApplicationArea = Suite;
                    Caption = 'Client Id';
                    ToolTip = 'Specifies the id of the Azure Active Directory application that will be used to connect to the SharePoint environment.', Comment = 'SharePoint and Azure Active Directory are names of a Microsoft service and a Microsoft Azure resource and should not be translated.';
                }
                field("Client Secret"; ClientSecret)
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Client Secret';
                    ExtendedDatatype = Masked;
                    ToolTip = 'Specifies the secret of the Azure Active Directory application that will be used to connect to the SharePoint environment.', Comment = 'SharePoint and Azure Active Directory are names of a Microsoft service and a Microsoft Azure resource and should not be translated.';

                    trigger OnValidate()
                    var
                        DocumentServiceManagement: Codeunit "Document Service 4PS";
                    begin
                        if not IsTemporary() then
                            if (ClientSecret <> '') and (not EncryptionEnabled()) then
                                if Confirm(EncryptionIsNotActivatedQst) then
                                    Page.RunModal(Page::"Data Encryption Management");
                        DocumentServiceManagement.SetClientSecret(ClientSecret);
                    end;
                }
                field("Redirect URL"; "Redirect URL")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the Redirect URL of the Azure Active Directory application that will be used to connect to the SharePoint environment.', Comment = 'SharePoint and Azure Active Directory are names of a Microsoft service and a Microsoft Azure resource and should not be translated.';
                }
            }
        }
        area(factboxes)
        {
            systempart(Control1100528408; Links)
            {
                Visible = false;
            }
            systempart(Control1100528407; Notes)
            {
                Visible = false;
            }
        }
    }

    actions
    {
        area(processing)
        {
            action("Test Connection")
            {
                Caption = 'Test Connection';
                Image = ValidateEmailLoggingSetup;
                Promoted = true;
                PromotedCategory = Process;
                Visible = IsLegacyAuthentication;
                ApplicationArea = Basic, Suite;

                trigger OnAction()
                var
                    DocumentServiceManagement: Codeunit "Document Service 4PS";
                begin
                    // Save record to make sure the credentials are reset.
                    Modify;
                    DocumentServiceManagement.SetUseDocumentService4PS(true);
                    DocumentServiceManagement.TestConnection;
                    Message(Text000);
                end;
            }
        }
    }

    trigger OnOpenPage()
    var
        DocumentService: Record "Document Service";
    begin
        // Reset;
        // if not Get then begin
        //     Init;
        //     "Authentication Type" := "Authentication Type"::OAuth2;
        //     InitializeDefaultRedirectUrl();
        //     IsLegacyAuthentication := false;
        //     Insert;
        // end else begin
        //     IsLegacyAuthentication := Rec."Authentication Type" = Rec."Authentication Type"::Legacy;
        //     if not IsLegacyAuthentication then begin
        //         ClientSecret := GetClientSecret();
        //         if "Redirect URL" = '' then
        //             InitializeDefaultRedirectUrl();
        //     end;
        //     Modify(false);
        // end;

        // if DocumentService.IsEmpty() then begin
        //     DocumentService.Init;
        //     DocumentService."Service ID" := 'Service 1';
        //     DocumentService."Authentication Type" := "Authentication Type"::OAuth2;
        //     DocumentService.Insert(false);
        // end;
    end;

    var
        Text000: Label 'The connection settings validated correctly, and the current configuration can connect to the document storage service.';
        EncryptionIsNotActivatedQst: Label 'Data encryption is currently not enabled. We recommend that you encrypt data. \Do you want to open the Data Encryption Management window?';
        IsLegacyAuthentication: Boolean;
        ClientSecret: Text;
        SoftwareAsAService: Boolean;

    local procedure InitializeDefaultRedirectUrl()
    var
        AzureADMgt: Codeunit "Azure AD Mgt.";
    begin
        "Redirect URL" := AzureADMgt.GetDefaultRedirectUrl();
    end;

    local procedure GetClientSecret(): Text
    var
        DocumentServiceManagement: Codeunit "Document Service 4PS";
        ClientSecretTxt: Text;
    begin
        if DocumentServiceManagement.TryGetClientSecretFromIsolatedStorage(ClientSecretTxt) then
            exit(ClientSecretTxt);

        exit('');
    end;
}

