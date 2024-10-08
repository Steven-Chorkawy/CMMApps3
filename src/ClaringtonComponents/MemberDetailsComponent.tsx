import { Text, Stack, Breadcrumb, IBreadcrumbItem } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import IMemberListItem from '../ClaringtonInterfaces/IMemberListItem';
import { ICommitteeMemberHistoryListItem } from '../ClaringtonInterfaces/INewCommitteeMemberHistoryListItem';
import { CalculateTotalYearsServed, GetMembersTermHistory } from '../HelperMethods/MyHelperMethods';
import { MyShimmer } from './MyShimmer';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

//#region 
export interface IMemberDetailsComponentProps {
    memberId: number;
    title?: string;      // Title of the component if any are required.
    context: WebPartContext | ListViewCommandSetContext;
}

export interface IMemberDetailsComponentState {
    member: any;
    allTermHistories: ICommitteeMemberHistoryListItem[];    // A list of all the members terms.  All terms from all committees.
    termHistories: ICommitteeMemberHistoryListItem[];       // A list of the members most recent term of each committee.  Only one term per committee.
}

export interface ICommitteeMemberBreadCrumbProps {
    context: WebPartContext | ListViewCommandSetContext;
    committeeTerm: ICommitteeMemberHistoryListItem;
    allTerms?: ICommitteeMemberHistoryListItem[];     // Used to preview past committees.
}

export interface ICommitteeMemberContactDetails {
    member: IMemberListItem;
}

export interface ICommitteeMemberTermHistoryProps {
    memberID: number;
    context: WebPartContext | ListViewCommandSetContext;
}

export interface ICommitteeMemberTermHistoryState {
    allTermHistories: any[];
    termHistories: any[];
}
//#endregion

export class CommitteeMemberBreadCrumb extends React.Component<ICommitteeMemberBreadCrumbProps, any> {
    constructor(props: any) {
        super(props);
    }

    public render(): React.ReactElement<any> {
        //const ID_FILTER = `?=FilterValue72&FilterField1=Member_x0020_Display_x0020_Name_x003a_ID&FilterValue1=${this.props.committeeTerm.MemberID}`;
        const LIBRARY_URL = `https://claringtonnet.sharepoint.com/sites/CMM/${this.props.committeeTerm.CommitteeName}`;


        const itemsWithHref: IBreadcrumbItem[] = [
            // Normally each breadcrumb would have a unique href, but to make the navigation less disruptive
            // in the example, it uses the breadcrumb page as the href for all the items
            {
                text: this.props.committeeTerm.CommitteeName,
                key: 'CommitteeLibrary',
                href: `${LIBRARY_URL}`,
                target: '_blank',
                title: `View all ${this.props.committeeTerm.CommitteeName} committee members`
            },
            {
                text: `${this.props.committeeTerm.Title}`,
                // key: 'Member', href: `${LIBRARY_URL}${ID_FILTER}`, isCurrentItem: true
                key: 'Member',
                href: `${LIBRARY_URL}/${this.props.committeeTerm.Title}`,
                isCurrentItem: true,
                target: '_blank',
                title: `View ${this.props.committeeTerm.Title} documents for ${this.props.committeeTerm.CommitteeName}`
            },
        ];

        return <div>
            <Breadcrumb
                items={itemsWithHref}
                maxDisplayedItems={2}
                ariaLabel="Breadcrumb with items rendered as buttons"
                overflowAriaLabel="More links"
            />
            <div>
                <div>
                    <Text variant={'small'}>
                        <span title={`Start Date`}>{new Date(this.props.committeeTerm.StartDate).toLocaleDateString()}</span> - <span title={`End Date`}>{new Date(this.props.committeeTerm.OData__EndDate).toLocaleDateString()}</span>
                    </Text>
                </div>
                {/* <ActivityItem {...activityItem} key={activityItem.key} className={classNames.exampleRoot} /> */}
            </div>
        </div >;
    }
}

export class CommitteeMemberTermHistory extends React.Component<ICommitteeMemberTermHistoryProps, ICommitteeMemberTermHistoryState> {
    constructor(props: any) {
        super(props);
        this.state = {
            allTermHistories: undefined,
            termHistories: undefined
        };

        GetMembersTermHistory(this.props.memberID)
            .then(values => {
                this.setState({
                    allTermHistories: values,
                    termHistories: values.filter((value, index, self) => index === self.sort((a, b) => {
                        // Turn your strings into dates, and then subtract them
                        // to get a value that is either negative, positive, or zero.
                        const bb: any = new Date(b.StartDate),
                            aa: any = new Date(b.StartDate);
                        // This sorts the term histories.
                        return bb - aa;
                    }).findIndex((t) => (t.CommitteeName === value.CommitteeName)))
                });
            })
            .catch(reason => {
                console.error('Failed ot get members term history!');
                console.error(reason);
            });
    }

    public render(): React.ReactElement<any> {
        return this.state.termHistories ?
            <div>
                {
                    this.state.allTermHistories &&
                    <div>
                        <div>
                            <Text variant="xLarge">{CalculateTotalYearsServed(this.state.allTermHistories)} Years Served.</Text>
                        </div>
                        <div>
                            <Text variant="small">{this.state.allTermHistories.length} Terms on {this.state.termHistories.length} Committees</Text>
                        </div>
                    </div>
                }
                {this.state.termHistories.map((term: any, index: number) => {
                    return <div key={`termHistory-${index}`}>
                        <CommitteeMemberBreadCrumb
                            committeeTerm={term}
                            allTerms={this.state.allTermHistories}
                            context={this.props.context} />
                    </div>;
                })}
            </div> :
            <MyShimmer />;
    }
}

export class CommitteeMemberContactDetails extends React.Component<ICommitteeMemberContactDetails, {}> {
    /**
     *
     */
    constructor(props: ICommitteeMemberContactDetails) {
        super(props);
    }

    private _detailDisplay = (prop: string, label: string): React.ReactElement<any> => {
        return <div><span>{label}: {this.props.member[prop as keyof IMemberListItem] && this.props.member[prop as keyof IMemberListItem]}</span></div>;
    }

    public render(): React.ReactElement<any> {
        return <div>
            <Stack horizontal={true} wrap={true}>
                <Stack.Item grow={6}>
                    {this._detailDisplay('EMail', 'Email')}
                    {this._detailDisplay('CellPhone1', 'Cell Phone')}
                    {this._detailDisplay('HomePhone', 'Home Phone')}
                </Stack.Item>
                <Stack.Item grow={6}>
                    {this._detailDisplay('WorkAddress', 'Address')}
                    {this._detailDisplay('PostalCode', 'Postal Code')}
                    {this._detailDisplay('WorkCity', 'City')}
                </Stack.Item>
            </Stack>
            <Stack horizontal={true} wrap={true} style={{ marginTop: '5px' }}>
                <Stack.Item grow={6}>
                    {this._detailDisplay('GenderChoice', 'Gender')}
                    {this._detailDisplay('Age', 'Age')}
                    {this._detailDisplay('Disability', 'Disability')}
                </Stack.Item>
                <Stack.Item grow={6}>
                    {this._detailDisplay('IdentifyIndigenous', 'Identify as Indigenous')}
                    {this._detailDisplay('EthnoCultural', 'Ethno-Cultural Identity')}
                    {this._detailDisplay('RacialBackground', 'Racial Background')}
                </Stack.Item>
            </Stack>
        </div>;
    }
}
