export default interface IMemberListItem {
    Id?: number;
    ID?: number;
    Title: string;
    FirstName: string;      // ! Required.
    // MiddleName?: string;
    LastName: string;       // ! Required.
    DisplayName?: string;
    // Salutation?: string;

    EMail?: string;
    // Email2?: string;
    CellPhone1?: string;
    // WorkPhone?: string;     // Display name Buesiness Phone.
    HomePhone?: string;

    WorkAddress?: string;   // Display name Address.
    // Birthday: string;       // This is a Date and Time in SharePoint. 
    WorkCity?: string;      // Display name City.
    WorkCountry?: string;   // Default to Canada in SharePoint.
    PostalCode?: string;
    Province?: string;      // This is a Choice column in SharePoint.
    RetentionDate?: string;
    RetentionDateCommittee?: string;

    EthnoCultural?: string; // Choice column in SharePoint.
    GenderChoice?: string; // Choice column in SharePoint.  Display name is 'Gender'.
    IdentifyIndigenous?: string; // Choice column in SharePoint.
    RacialBackground?: string; // Choice column in SharePoint.
    Age?: string; // Choice column in SharePoint.
    Disability?: string; // Choice column in SharePoint.

    Modified?: Date;
}