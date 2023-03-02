page 50004 "Sharepoint Connector List"
{
    PageType = ListPart;
    SourceTable = "Sharepoint Connector List";
    Editable = false;

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field(Name; Rec.Title)
                {
                    ApplicationArea = All;
                }
                field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                }
                field(OdataId; Rec.OdataId)
                {
                    ApplicationArea = All;
                }
                field("Server Relative Url"; Rec."Server Relative Url")
                {
                    ApplicationArea = All;
                }
            }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.DeleteAll();
    end;
}