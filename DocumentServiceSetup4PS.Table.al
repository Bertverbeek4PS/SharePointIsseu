table 50100 "Document Service Setup 4PS"
{
    Caption = 'Document Service Setup 4PS';
    Permissions = TableData 2000000114 = rimd;

    fields
    {
        field(10; "Primary Key"; Code[10])
        {
            Caption = 'Primary Key';
        }
        field(15; "Client Id"; Text[250])
        {
            Caption = 'Client Id';
        }
        field(19; "Redirect URL"; Text[2048])
        {
            Caption = 'Redirect URL';
        }
        field(20; "SharePoint Site URL"; Text[250])
        {
            Caption = 'SharePoint Site URL';

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService.Location := "SharePoint Site URL";
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(21; "Authentication Type"; Option)
        {
            Caption = 'Authentication Type';
            OptionCaption = 'Legacy,OAuth2';
            OptionMembers = Legacy,OAuth2;

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService."Authentication Type" := "Authentication Type";
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(23; "Client Secret Key"; Guid)
        {
            Caption = 'Client Secret Key';
        }
        field(30; "Document Library"; Text[250])
        {
            Caption = 'Document Library';

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService."Document Repository" := "Document Library";
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(40; Folder; Text[250])
        {
            Caption = 'Folder';

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService.Folder := Folder;
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(50; "User Name"; Text[128])
        {
            Caption = 'User Name';

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService."User Name" := "User Name";
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(60; Password; Text[128])
        {
            Caption = 'Password';
            ExtendedDatatype = Masked;

            trigger OnValidate()
            var
                DocumentService: Record "Document Service";
            begin
                // if DocumentService.FindFirst then begin
                //     DocumentService.Password := Password;
                //     DocumentService.Modify;
                // end;
            end;
        }
        field(70; "Default Ext. Doc. Subdir."; Text[250])
        {
            Caption = 'Default Ext. Doc. Subdir.';
            DataClassification = ToBeClassified;
        }
        field(80; "Subdir. Document Parts"; Text[250])
        {
            Caption = 'Subdir. Document Parts';
            DataClassification = ToBeClassified;

            trigger OnValidate()
            var
                FileManagement: Codeunit "File Management";
            begin
                //FileManagement.AddBackSlashToDirectoryName("Subdir. Document Parts");
            end;
        }
    }

    keys
    {
        key(Key1; "Primary Key")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }
}

