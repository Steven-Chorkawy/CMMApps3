import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/security";
import { IContentTypeInfo } from "@pnp/sp/content-types";
import ICommitteeFileItem from "../ClaringtonInterfaces/ICommitteeFileItem";
import { MyLists } from "./MyLists";
import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import IMemberListItem from "../ClaringtonInterfaces/IMemberListItem";
import { IFolderAddResult } from "@pnp/sp/folders";
import { INewCommitteeMemberHistoryListItem, ICommitteeMemberHistoryListItem } from "../ClaringtonInterfaces/INewCommitteeMemberHistoryListItem";
import { ListViewCommandSetContext, RowAccessor } from "@microsoft/sp-listview-extensibility";

let _sp: SPFI = null;

export const getSP = (context?: WebPartContext | ListViewCommandSetContext): SPFI => {
    if (_sp === null && context !== null) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};

//#region Constants
export const FORM_DATA_INDEX = "formDataIndex";

// Content Type ID of the Document Set found in the Committees Document Library. 
export const COMMITTEE_FILE_CONTENT_TYPE_ID = "0x0120D5200038D10D0D1AF55A4DB6F57F794DB8B0CD";
//#endregion

//#region
export const CONSOLE_LOG_ERROR = (reason: any, customMessage?: string): void => {
    console.error(customMessage ? customMessage : "Something went wrong!");
    console.error(reason);
};
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
    const library = await sp.web.lists.getByTitle(libraryTitle).select('Title', 'RootFolder/ServerRelativeUrl').expand('RootFolder')();
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
export const CalculateMemberInfoRetention = async (memberId: number): Promise<{ date: Date, committee: string }> => {
    const sp = getSP();
    let output: Date = undefined;
    let committeeName: string = undefined;
    const RETENTION_PERIOD = 5; // Retention is 5 years + last Term End Date.

    const memberHistory = await sp.web.lists.getByTitle(MyLists.CommitteeMemberHistory).items
        .filter(`SPFX_CommitteeMemberDisplayNameId eq ${memberId}`)
        .orderBy('OData__EndDate', false)();

    if (memberHistory && memberHistory.length > 0) {
        const tmpDate = new Date(memberHistory[0].OData__EndDate);
        output = new Date(tmpDate.getFullYear() + RETENTION_PERIOD, tmpDate.getMonth(), tmpDate.getDate());
        committeeName = memberHistory[0].CommitteeName;
    }

    return { date: output, committee: committeeName };
};

export const CalculateTotalYearsServed = (committeeTerms: ICommitteeMemberHistoryListItem[]): number => {
    /**
     * Steps to confirm Total Years Served.
     * 1.   Start date must be less than today.  If is not ignore this term as it is invalid.
     * 2.   End date must be greater than or equal to day.  If it is not use today's date.
     * 3.   
     */
    let totalYears: number = 0;
    let termTotal: number = 0;

    for (let termIndex = 0; termIndex < committeeTerms.length; termIndex++) {
        // reset this counter. 
        termTotal = 0;

        const TERM = committeeTerms[termIndex];
        const START_DATE = new Date(TERM.StartDate);
        const TODAY = new Date();
        let endDate = new Date(TERM.OData__EndDate);


        if (START_DATE > TODAY) {
            continue; // Continue onto the next iteration. 
        }

        // End date is currently in the future so we will use today's date to calculate the total terms served. 
        if (endDate >= TODAY) {
            endDate = TODAY;
        }

        termTotal = endDate.getFullYear() - START_DATE.getFullYear();

        // Add to the running total.
        totalYears += termTotal;
    }

    return totalYears;
};
//#endregion

//#region Reads
export const GetChoiceColumn = async (listTitle: string, columnName: string): Promise<string[]> => {
    const sp = getSP();
    try {
        const choiceColumn: any = await sp.web.lists.getByTitle(listTitle).fields.getByTitle(columnName).select('Choices')();
        return choiceColumn.Choices;
    } catch (error) {
        console.error('Something went wrong in GetChoiceColumn!');
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
        const output: any = await sp.web.lists.getByTitle(MyLists.CommitteeDocuments).items.filter(`Title eq '${committeeName}'`)();

        if (output && output.length === 1) {
            return output[0];
        }
        else {
            throw Error(`Multiple '${committeeName}' found!`);
        }
    } catch (error) {
        console.error('Something went wrong in GetChoiceColumn!');
        console.error(error);
        return undefined;
    }
};

export const GetListOfActiveCommittees = async (): Promise<any> => {
    const sp = getSP();
    return await sp.web.lists.getByTitle(MyLists.CommitteeDocuments).items.filter(`OData__Status eq 'Active' and ContentTypeId eq '${COMMITTEE_FILE_CONTENT_TYPE_ID}'`)();
};

export const GetLibraryContentTypes = async (libraryTitle: string): Promise<string> => {
    const sp = getSP();
    const library = await sp.web.lists.getByTitle(libraryTitle);
    return (await library.contentTypes()).find((f: IContentTypeInfo) => f.Group === "Custom Content Types" && f.StringId.includes('0x0120')).StringId;

};

export const GetMembers = async (): Promise<IMemberListItem[]> => await getSP().web.lists.getByTitle(MyLists.Members).items();

export const GetMember = async (id: number): Promise<any> => await getSP().web.lists.getByTitle(MyLists.Members).items.getById(id)();

/**
 * Get a members term history.
 * @param id MemberID field from the Committee Member History list.
 * @returns ICommitteeMemberHistoryListItem[]
 */
export const GetMembersTermHistory = async (id: number): Promise<ICommitteeMemberHistoryListItem[]> => await getSP().web.lists.getByTitle(MyLists.CommitteeMemberHistory).items.filter(`MemberID eq ${id}`)();

/**
 * Get the Member ID from a given folder path.
 * This method works for Lists and Document Libraries.
 * @param fileRef Path to a folder/ Document Set.
 * @returns ID of the member in the Members list.
 */
export const GetMemberIdFromFileRef = async (fileRef: string): Promise<number> => {
    let output = NaN;
    const sp = getSP();
    const itemMetadata = await (await sp.web.getFolderByServerRelativePath(fileRef).getItem())();
    // let output11 = await output1();
    if (itemMetadata.MemberLookupId) {
        output = itemMetadata.MemberLookupId;
    }
    else if (itemMetadata.MemberID) {
        output = itemMetadata.MemberID;
    }
    return output;
}

/**
 * Parse the members ID of a given row from a CommandSet button click event.
 * @param selectedRow Selected row from IListViewCommandSetExecuteEventParameters event.
 * @returns ID in the form of a number.
 */
export const GetMemberIdFromSelectedRow = async (selectedRow: RowAccessor): Promise<number> => {
    const fileRef = selectedRow.getValueByName('FileRef');

    // First check if this list is the Members list.
    // If this is the Members list then all we need is the ID.  
    // * Note: that this is the only list that we can use the ID from because it is the Members list.
    return fileRef.includes('/Lists/Members') ? selectedRow.getValueByName('ID') : await GetMemberIdFromFileRef(fileRef);
}
//#endregion

//#region Create
/**
 * Create a new list item in the Committee Member History list.
 */
export const CreateCommitteeMemberHistoryItem = async (committeeMemberHistoryItem: INewCommitteeMemberHistoryListItem): Promise<void> => {
    const sp = getSP();
    await sp.web.lists.getByTitle(MyLists.CommitteeMemberHistory).items.add({ ...committeeMemberHistoryItem });

    // ? Why did I have this? 
    //const committeeMemberContactInfoRetention = await CalculateMemberInfoRetention(committeeMemberHistoryItem.SPFX_CommitteeMemberDisplayNameId);

    // ? What does this do?
    // await sp.web.lists.getByTitle(MyLists.Members).items.getById(committeeMemberHistoryItem.SPFX_CommitteeMemberDisplayNameId).update({
    //     RetentionDate: committeeMemberContactInfoRetention.date,
    //     RetentionDateCommittee: committeeMemberContactInfoRetention.committee
    // });
};

export const CreateNewMember = async (member: IMemberListItem): Promise<IItemAddResult> => {
    const sp = getSP();

    member.Title = FormatMemberTitle(member.FirstName, member.LastName);
    // add an item to the list
    return await sp.web.lists.getByTitle(MyLists.Members).items.add(member);
};

export const CreateDocumentSet = async (input: any): Promise<IItemUpdateResult> => {
    let newFolderResult: IFolderAddResult;
    const FOLDER_NAME = await FormatDocumentSetPath(input.LibraryTitle, input.Title);
    let libraryDocumentSetContentTypeId;
    const sp = getSP();

    try {
        libraryDocumentSetContentTypeId = await GetLibraryContentTypes(input.LibraryTitle);
        if (!libraryDocumentSetContentTypeId) {
            throw Error("Error! Cannot get content type for library.");
        }

        // Because sp.web.folders.add overwrites existing folder I have to do a manual check.
        if (await CheckForExistingDocumentSetByServerRelativePath(FOLDER_NAME)) {
            throw new Error(`Error! Cannot Create new Document Set. Duplicate Name detected. "${FOLDER_NAME}"`);
        }

        newFolderResult = await sp.web.folders.addUsingPath(FOLDER_NAME);
    } catch (error) {
        console.error(error);
        throw error;
    }

    const newFolderProperties = await sp.web.getFolderByServerRelativePath(newFolderResult.data.ServerRelativeUrl).listItemAllFields();
    return await sp.web.lists.getByTitle(input.LibraryTitle).items.getById(newFolderProperties.ID).update({
        ContentTypeId: libraryDocumentSetContentTypeId
    });
};

/**
 * Create a document set for an existing member in a committee library.
 * @param member ID of the member to add to a committee.
 * @param committee Committee to add member to.
 * TODO: What type should the committee param be?
 */
export const CreateNewCommitteeMember = async (memberId: number, committee: any): Promise<void> => {
    const sp = getSP();
    if (!committee) {
        throw Error("No Committee provided.");
    }

    const member = await sp.web.lists.getByTitle(MyLists.Members).items.getById(memberId)();
    const PATH_TO_DOC_SET = await FormatDocumentSetPath(committee.CommitteeName, member.Title);

    // Step 1: Create the document set.
    const docSet = await (await CreateDocumentSet({ LibraryTitle: committee.CommitteeName, Title: member.Title })).item();

    // Step 2: Update Metadata.
    await sp.web.lists.getByTitle(committee.CommitteeName).items.getById(docSet.ID).update({
        Position: committee.Position,
        OData__Status: committee._Status,
        OData__EndDate: committee._EndDate,
        StartDate: committee.StartDate,
        MemberLookupId:memberId
    });

    // Step 3: Upload Attachments. 
    if (committee.Files) {
        committee.Files.map((file: any) => {
            file.downloadFileContent().then((fileContent: any) => {
                sp.web.getFolderByServerRelativePath(PATH_TO_DOC_SET).files.addUsingPath(file.fileName, fileContent, { Overwrite: true }).catch(reason => {
                    console.error('Failed to upload attachment');
                    console.error(reason);
                });
            });
        });
    }

    // Step 4: Update Committee Member List Item to include this new committee.
    // TODO: How do I manage this relationship? 

    // Step 5: Create a committee member history list item record.
    await CreateCommitteeMemberHistoryItem({
        CommitteeName: committee.CommitteeName,
        Title: FormatMemberTitle(member.FirstName, member.LastName),
        OData__EndDate: committee._EndDate,
        StartDate: committee.StartDate,
        MemberLookupId: memberId,
        MemberID: memberId
    });
};

/**
 * Renew a Committee Member by updating their Status, Position, Start Date, End Date, and upload new attachments in the Committee library.
 * Also create a new list item in the Committee Member History list.  
 * @param memberId ID from the Members list.
 * @param committeeMemberProperties All properties and fields from the Renew Committee Member form.
 */
export const RenewCommitteeMember = async (memberId: number, committeeMemberProperties: any): Promise<void> => {
    console.log('RenewCommitteeMember started...');
    const sp = getSP();
    const committeeLibrary = sp.web.lists.getByTitle(committeeMemberProperties.committeeName);

    // * Step 1: Get the Document set for the current Committee.
    let committeeMemberDocumentSet = await committeeLibrary.getItemsByCAMLQuery({
        ViewXml: `
        <View>
            <Query>
                <Where>
                <Eq>
                    <FieldRef Name="MemberLookup"/>
                    <Value Type="Lookup">${memberId}</Value>
                </Eq>
                </Where>
            </Query>
            <RowLimit>1</RowLimit>
        </View>
        `,
    });

    // If we have anything other than 1 result, something went wrong.
    if (committeeMemberDocumentSet.length !== 1) {
        throw "Something went wrong while querying Committee Member Document Set...";
    }

    committeeMemberDocumentSet = committeeMemberDocumentSet[0];
    // IDK why but getItemsByCAMLQuery() cannot get FileLeafRef for some reason!!
    const documentSetTitle = await committeeLibrary.items.getById(committeeMemberDocumentSet.ID).select('FileLeafRef')();
    committeeMemberDocumentSet.Title = documentSetTitle.FileLeafRef;

    // * Step 2: Update the Doc Sets Status, Position, Start Date, and End Date.
    const committeeMemberDocumentSet_UpdateResult = await committeeLibrary.items.getById(committeeMemberDocumentSet.ID).update({
        OData__Status: committeeMemberProperties._Status,
        Position: committeeMemberProperties.Position,
        StartDate: committeeMemberProperties.StartDate,
        OData__EndDate: committeeMemberProperties._EndDate
    });

    // * Step 3: Upload any attachments to the Doc Set.
    if (committeeMemberProperties.Files) {
        const PATH_TO_DOC_SET = await FormatDocumentSetPath(committeeMemberProperties.committeeName, committeeMemberDocumentSet.Title);
        committeeMemberProperties.Files.map((file: any) => {
            file.downloadFileContent().then((fileContent: any) => {
                sp.web.getFolderByServerRelativePath(PATH_TO_DOC_SET).files.addUsingPath(file.fileName, fileContent, { Overwrite: true }).catch(reason => {
                    console.error('Failed to upload attachment');
                    console.error(reason);
                });
            });
        });
    }

    // * Step 4: Update Committee Member History.
    // This is an async method but we really don't need to wait for the results.
    CreateCommitteeMemberHistoryItem({
        Title: committeeMemberDocumentSet.Title,
        CommitteeName: committeeMemberProperties.committeeName,
        OData__EndDate: new Date(committeeMemberProperties._EndDate),
        StartDate: new Date(committeeMemberProperties.StartDate),
        MemberID: memberId,
        MemberLookupId: memberId
    });

    // * Step 5: ... TBD Update something else?...
}
//#endregion
