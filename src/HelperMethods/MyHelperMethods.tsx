import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/security";
import ICommitteeFileItem from "../ClaringtonInterfaces/ICommitteeFileItem";
import { MyLists } from "./MyLists";
import { IItemAddResult, IItems } from "@pnp/sp/items";
import IMemberListItem from "../ClaringtonInterfaces/IMemberListItem";

let _sp: SPFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
    if (_sp === null && context !== null) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};

//#region Constants
export const FORM_DATA_INDEX = "formDataIndex";
export const COMMITTEE_FILE_CONTENT_TYPE_ID = "0x0120D52000DA0D74B289281E4CBF23681415CE96AF";
//#endregion


//#region Formatters
/**
 * Format Fluent UI DatePicker.
 * @param date Date input from Fluent UI DatePicker
 * @returns Month/Day/Year as a string.
 */
export const OnFormatDate = (date?: Date): string => {
    return !date ? '' : (date.getMonth() + 1) + '/' + date.getDate() + '/' + (date.getFullYear());
};

/**
 * Format a path to a document set that will be created.
 * @param libraryTitle Title of Library
 * @param title Title of new Folder/Document Set to be created.
 * @returns Path to Document set as string.
 */
export const FormatDocumentSetPath = async (libraryTitle: string, title: string): Promise<string> => {
    const sp = getSP();
    let library = await sp.web.lists.getByTitle(libraryTitle).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
    return `${library.RootFolder.ServerRelativeUrl}/${title}`;
};

export const CheckForExistingDocumentSetByServerRelativePath = async (serverRelativePath: string): Promise<boolean> => {
    const sp = getSP();
    return await (await sp.web.getFolderByServerRelativePath(serverRelativePath).select('Exists')()).Exists;
};

/**
 * Calculate a term end date.
 * Term End Date = start date + Term Length.
 */
export const CalculateTermEndDate = (startDate: Date, termLength: number): Date => {
    return new Date(startDate.getFullYear() + termLength, startDate.getMonth(), startDate.getDate());
};

/**
 * Format members Full Name/ Title.
 * @param firstName Members First Name.
 * @param lastName Members Last Name.
 * @returns "lastName, firstName"
 */
 export const FormatMemberTitle = (firstName: string, lastName: string): string => { return `${lastName}, ${firstName}`; };

/**
 * Calculate a committee members personal contact information retention period.
 * Personal Contact Information retention period = last committee term end date + 5 years.
 * @param memberId ID of the member that we are trying to calculate for.
 * @returns The date the members personal contact info should be deleted.
 */
// export const CalculateMemberInfoRetention = async (memberId: number): Promise<{ date, committee }> => {
//     let output: Date = undefined;
//     let committeeName: string = undefined;
//     const RETENTION_PERIOD = 5; // Retention is 5 years + last Term End Date.
//     let memberHistory = await sp.web.lists.getByTitle(MyLists.CommitteeMemberHistory).items
//         .filter(`SPFX_CommitteeMemberDisplayNameId eq ${memberId}`)
//         .orderBy('TermEndDate', false).get();

//     if (memberHistory && memberHistory.length > 0) {
//         let tmpDate = new Date(memberHistory[0].TermEndDate);
//         output = new Date(tmpDate.getFullYear() + RETENTION_PERIOD, tmpDate.getMonth(), tmpDate.getDate());
//         committeeName = memberHistory[0].CommitteeName;
//     }

//     return { date: output, committee: committeeName };
// };

// export const CalculateTotalYearsServed = (committeeTerms: ICommitteeMemberHistoryListItem[]): number => {
//     /**
//      * Steps to confirm Total Years Served.
//      * 1.   Start date must be less than today.  If is not ignore this term as it is invalid.
//      * 2.   End date must be greater than or equal to day.  If it is not use today's date.
//      * 3.   
//      */
//     debugger;
//     let totalYears: number = 0;
//     let termTotal: number = 0;

//     for (let termIndex = 0; termIndex < committeeTerms.length; termIndex++) {
//         // reset this counter. 
//         termTotal = 0;

//         const term = committeeTerms[termIndex];
//         let startDate = new Date(term.StartDate),
//             endDate = new Date(term.TermEndDate),
//             today = new Date();

//         console.log(term);
//         if (startDate > today) {
//             debugger;
//             console.log('Something went wrong!');
//             continue; // Continue onto the next iteration. 
//         }

//         // End date is currently in the future so we will use today's date to calculate the total terms served. 
//         if (endDate >= today) {
//             endDate = today;
//         }

//         termTotal = endDate.getFullYear() - startDate.getFullYear();

//         // Add to the running total.
//         totalYears += termTotal;
//     }

//     return totalYears;
// };
//#endregion


//#region Read
export const GetChoiceColumn = async (listTitle: string, columnName: string): Promise<string[]> => {
    const sp = getSP();
    try {
        let choiceColumn: any = await sp.web.lists.getByTitle(listTitle).fields.getByTitle(columnName).select('Choices')();
        return choiceColumn.Choices;
    } catch (error) {
        console.log('Something went wrong in GetChoiceColumn!');
        console.error(error);
        return [];
    }
};

/**
 * Get committee data from the CommitteeFiles library.
 * @param committeeName Name of the Committee Document Set.
 * @returns CommitteeFiles Document Set metadata. 
 */
export const GetCommitteeByName = async (committeeName: string): Promise<ICommitteeFileItem> => {
    const sp = getSP();
    try {
        let output: any = await sp.web.lists.getByTitle(MyLists.CommitteeDocuments).items.filter(`Title eq '${committeeName}'`)();

        if (output && output.length === 1) {
            return output[0];
        }
        else {
            throw `Multiple '${committeeName}' found!`;
        }
    } catch (error) {
        console.log('Something went wrong in GetChoiceColumn!');
        console.error(error);
        return undefined;
    }
};

export const GetListOfActiveCommittees = async (): Promise<any> => {
    // TODO: Remove hard coded content type id.
    const sp = getSP();
    let output = await sp.web.lists.getByTitle(MyLists.CommitteeDocuments).items.filter(`OData__Status eq 'Active' and ContentTypeId eq '${COMMITTEE_FILE_CONTENT_TYPE_ID}'`)();
    return output;
};


// export const GetLibraryContentTypes = async (libraryTitle: string): Promise<string> => {
//     const sp = getSP();
//     let library = await sp.web.lists.getByTitle(libraryTitle);
//     return (await library.contentTypes()).find((f: IContentTypeInfo) => f.Group === "Custom Content Types" && f.StringId.includes('0x0120')).StringId;
// };

// export const GetMembers = async (): Promise<IMemberListItem[]> => await sp.web.lists.getByTitle(MyLists.Members).items.getAll();

export const GetMember = async (id: number): Promise<any> => await getSP().web.lists.getByTitle(MyLists.Members).items.getById(id);

/**
 * Get a members term history.
 * @param id MemberID field from the Committee Member History list.
 * @returns ICommitteeMemberHistoryListItem[]
 */
// export const GetMembersTermHistory = async (id: number): Promise<ICommitteeMemberHistoryListItem[]> => await sp.web.lists.getByTitle(MyLists.CommitteeMemberHistory).items.filter(`MemberID eq ${id}`).get();

//#region Create
export const CreateNewMember = async (member: IMemberListItem): Promise<IItemAddResult> => {
    const sp = getSP();

    member.Title = FormatMemberTitle(member.FirstName, member.LastName);
    // add an item to the list
    let iar = await sp.web.lists.getByTitle(MyLists.Members).items.add(member);
    return iar;
};
//#endregion

//#endregion