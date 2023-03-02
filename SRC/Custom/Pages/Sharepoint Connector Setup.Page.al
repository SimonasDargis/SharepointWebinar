page 50000 "Sharepoint Connector Setup"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "Sharepoint Connector Setup";
    ModifyAllowed = true;
    InsertAllowed = false;
    DeleteAllowed = false;

    layout
    {
        area(Content)
        {
            group(Setup)
            {
                field("Client ID"; Rec."Client ID")
                {
                    ApplicationArea = All;
                }
                field("Client Secret"; Rec."Client Secret")
                {
                    ApplicationArea = All;
                    ExtendedDatatype = Masked;
                }
                field("Sharepoint URL"; Rec."Sharepoint URL")
                {
                    ApplicationArea = All;
                }
            }
            group("File")
            {
                field(FileName; FileName)
                {
                    ApplicationArea = All;
                    Caption = 'File Name';
                }
                field(FileContents; FileContents)
                {
                    ApplicationArea = All;
                    Caption = 'File Contents';
                    MultiLine = true;
                }
                field(FileDirectory; FileDirectory)
                {
                    ApplicationArea = All;
                    Caption = 'Check for file';
                }
            }
            part(SCList; "Sharepoint Connector List")
            {
            }
        }
    }

    actions
    {
        area(Processing)
        {
            //CONNECT TO A SITE
            action(Connect)
            {
                ApplicationArea = All;
                Caption = 'Connect to a site';

                trigger OnAction()
                var
                    AadTenantId: Text;
                    Diag: Interface "HTTP Diagnostics";
                    SharePointList: Record "SharePoint List" temporary;
                begin
                    AadTenantId := GetAadTenantNameFromBaseUrl(Rec."Sharepoint URL"); //Used to get an Azure Active Directory ID from a URL
                    SharePointClient.Initialize(Rec."Sharepoint URL", GetSharePointAuthorization(AadTenantId)); //Initializes the client
                    SharePointClient.GetLists(SharePointList); //We need to perform at least one action to get diagnostics data
                    //Otherwise, the GetDiagnostics function will just return 0
                    Diag := SharePointClient.GetDiagnostics(); //Optional: used to get diagnostics, useful for debugging errors
                    Message('Diagnostics:\%1\\%2', Diag.GetHttpStatusCode(), Diag.GetResponseReasonPhrase());
                end;
            }
            action(GetList)
            {
                ApplicationArea = All;
                Caption = 'Get Sharepoint List';

                trigger OnAction()
                var
                    SaveFailedErr: Label 'Sharepoint list get failed.\ErrorMessage: %1\HttpRetryAfter: %2\HttpStatusCode: %3\ResponseReasonPhrase: %4', Comment = '%1=GetErrorMessage; %2=GetHttpRetryAfter; %3=GetHttpStatusCode; %4=GetResponseReasonPhrase';
                    AadTenantId: Text;
                    Diag: Interface "HTTP Diagnostics";
                    SharePointList: Record "SharePoint List" temporary;
                    //Used to display a Sharpoint list
                    SCListRec: Record "Sharepoint Connector List";
                    SharePointFolder: Record "SharePoint Folder";
                    SharePointFile: Record "SharePoint File";
                begin
                    if SharePointClient.GetLists(SharePointList) then begin //Sharepoint List is empty, GetLists writes data
                        SCListRec.DeleteAll(); //List is not a temporary record, delete to clear any data.
                        SharePointList.SetRange(Title, 'Documents'); //We filter out the Documents list to get a list of files.
                        if SharePointList.FindFirst() then begin
                            //We then proceed the folder of the Documents list so that we may get a list of files within the "Documents" directory
                            if SharePointClient.GetDocumentLibraryRootFolder(SharePointList.OdataId, SharePointFolder) then begin
                                //We then get the files from the root directory
                                //Note: We only get the list of files and not folders, this can be achieved by using
                                // GetSubFoldersByServerRelativeUrl after this one.
                                if SharePointClient.GetFolderFilesByServerRelativeUrl(SharePointFolder."Server Relative Url", SharePointFile) then begin
                                    SharePointFile.FindSet(); //Findset used to reset the record view since the repeat will only get the last record.
                                    repeat
                                        //Let's create a list of all items within the Documents directory
                                        //This list is used to display the files.
                                        SCListRec.Init();
                                        SCListRec.Id := SharePointFile."Unique Id";
                                        SCListRec.Title := SharePointFile.Name;
                                        SCListRec.OdataId := SharePointFile.OdataId;
                                        SCListRec."Server Relative Url" := SharePointFile."Server Relative Url";
                                        if SCListRec.Insert() then; //We initialize and insert a record for each Sharepoint List
                                    until SharePointFile.Next() = 0;
                                end;
                            end;
                        end;
                        CurrPage.SCList.Page.Update(); //Update the current page to refresh the list
                    end else begin
                        Diag := SharePointClient.GetDiagnostics(); //Get Diagnostics if error
                        Error(SaveFailedErr, Diag.GetErrorMessage(), Diag.GetHttpRetryAfter(), Diag.GetHttpStatusCode(), Diag.GetResponseReasonPhrase());
                    end;
                end;
            }
            action(ShowDirectory)
            {
                ApplicationArea = All;
                Caption = 'Show Directory';
                trigger OnAction()
                var
                    SaveFailedErr: Label 'Sharepoint list get failed.\ErrorMessage: %1\HttpRetryAfter: %2\HttpStatusCode: %3\ResponseReasonPhrase: %4', Comment = '%1=GetErrorMessage; %2=GetHttpRetryAfter; %3=GetHttpStatusCode; %4=GetResponseReasonPhrase';
                    AadTenantId: Text;
                    Diag: Interface "HTTP Diagnostics";
                    SharePointList: Record "SharePoint List" temporary;
                    SCListRec: Record "Sharepoint Connector List";
                    SharePointFolder: Record "SharePoint Folder";
                    SharePointFile: Record "SharePoint File";
                    filename: Text;
                    strpos: Integer;
                begin
                    SCListRec.DeleteAll();
                    if not SharePointClient.GetSubFoldersByServerRelativeUrl(FileDirectory, SharePointFolder) then
                        Error('Directory not found.');

                    if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, SharePointFile) then
                        Error('Directory not found.');

                    if SharePointFolder.FindSet() then begin
                        repeat
                            //Let's create a list of all items within the Documents directory
                            //This list is used to display the folders.
                            SCListRec.Init();
                            SCListRec.Id := SharePointFolder."Unique Id";
                            SCListRec.Title := SharePointFolder.Name;
                            SCListRec.OdataId := SharePointFolder.OdataId;
                            SCListRec."Server Relative Url" := SharePointFolder."Server Relative Url";
                            if SCListRec.Insert() then; //We initialize and insert a record for each folder
                        until SharePointFolder.Next() = 0;
                    end;

                    if SharePointFile.FindSet() then begin
                        repeat
                            //Let's create a list of all items within the Documents directory
                            //This list is used to display the files.
                            SCListRec.Init();
                            SCListRec.Id := SharePointFile."Unique Id";
                            SCListRec.Title := SharePointFile.Name;
                            SCListRec.OdataId := SharePointFile.OdataId;
                            SCListRec."Server Relative Url" := SharePointFile."Server Relative Url";
                            if SCListRec.Insert() then; //We initialize and insert a record for each file
                        until SharePointFile.Next() = 0;
                    end;

                    CurrPage.SCList.Page.Update(); //Update the current page to refresh the list
                end;
            }
            action(CreateFile)
            {
                ApplicationArea = All;
                Caption = 'Create File';

                trigger OnAction()
                var
                    SharePointFile: Record "SharePoint File" temporary;
                    SharePointList: Record "SharePoint List" temporary;
                    SharePointFolder: Record "SharePoint Folder" temporary;
                    SaveFailedErr: Label 'Save to SharePoint failed.\ErrorMessage: %1\HttpRetryAfter: %2\HttpStatusCode: %3\ResponseReasonPhrase: %4', Comment = '%1=GetErrorMessage; %2=GetHttpRetryAfter; %3=GetHttpStatusCode; %4=GetResponseReasonPhrase';
                    AadTenantId: Text;
                    IS: InStream;
                    OS: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    Diag: Interface "HTTP Diagnostics";
                begin
                    OS := TempBlob.CreateOutStream();
                    OS.Write(FileContents);
                    IS := TempBlob.CreateInStream();

                    if SharePointClient.GetLists(SharePointList) then begin
                        SharePointList.SetRange(Title, 'Documents');
                        if SharePointList.FindFirst() then begin
                            if SharePointClient.GetDocumentLibraryRootFolder(SharePointList.OdataId, SharePointFolder) then //Can use Id instead of title
                                SharePointClient.AddFileToFolder(FileDirectory, FileName, IS, SharePointFile);
                        end;
                    end else begin
                        Diag := SharePointClient.GetDiagnostics();
                        Error(SaveFailedErr, Diag.GetErrorMessage(), Diag.GetHttpRetryAfter(), Diag.GetHttpStatusCode(), Diag.GetResponseReasonPhrase());
                    end;
                end;
            }
            action(DownloadFile)
            {
                ApplicationArea = All;
                Caption = 'Download Selected File';

                trigger OnAction()
                var
                    SCListRec: Record "Sharepoint Connector List";
                begin
                    CurrPage.SCList.Page.GetRecord(SCListRec);
                    SharePointClient.DownloadFileContent(SCListRec.OdataId, SCListRec.Title);
                end;
            }
        }
    }

    local procedure GetSharePointAuthorization(AadTenantId: Text): Interface "SharePoint Authorization"
    var
        SharePointAuth: Codeunit "SharePoint Auth.";
        Scopes: List of [Text];
    begin
        Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default'); //Using a default scope provided as an example
        //We return an authorization code that will be used to initialize the Sharepoint Client
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, Rec."Client ID", Rec."Client Secret", Scopes));
    end;

    local procedure GetAadTenantNameFromBaseUrl(BaseUrl: Text): Text
    var
        Uri: Codeunit Uri;
        MySiteHostSuffixTxt: Label '-my.sharepoint.com', Locked = true;
        SharePointHostSuffixTxt: Label '.sharepoint.com', Locked = true;
        OnMicrosoftTxt: Label '.onmicrosoft.com', Locked = true;
        UrlInvalidErr: Label 'The Base Url %1 does not seem to be a valid SharePoint Online Url.', Comment = '%1=BaseUrl';
        Host: Text;
    begin
        // SharePoint Online format:  https://tenantname.sharepoint.com/SiteName/LibraryName/
        // SharePoint My Site format: https://tenantname-my.sharepoint.com/personal/user_name/
        Uri.Init(BaseUrl);
        Host := Uri.GetHost();
        if not Host.EndsWith(SharePointHostSuffixTxt) then
            Error(UrlInvalidErr, BaseUrl);
        if Host.EndsWith(MySiteHostSuffixTxt) then
            exit(CopyStr(Host, 1, StrPos(Host, MySiteHostSuffixTxt) - 1) + OnMicrosoftTxt);
        exit(CopyStr(Host, 1, StrPos(Host, SharePointHostSuffixTxt) - 1) + OnMicrosoftTxt);
    end;

    trigger OnOpenPage()
    begin
        if not Rec.Get() then begin
            Rec.Init();
            Rec.Insert();
        end;
    end;

    var
        FileName: Text;
        FileContents: Text;
        SharePointClient: Codeunit "SharePoint Client";
        FileDirectory: Text;

}