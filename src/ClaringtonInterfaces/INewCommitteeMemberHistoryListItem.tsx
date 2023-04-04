export default interface INewCommitteeMemberHistoryListItem {
    Title: string;
    CommitteeName: string;
    OData__EndDate: Date; // ? Why is SharePoint adding OData infront of this column name?
    StartDate: Date;
    FirstName: string;
    LastName: string;
    MemberID: number;
}