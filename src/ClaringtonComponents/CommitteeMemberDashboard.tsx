import * as React from 'react';
import { ComboBox, Icon } from '@fluentui/react';
import { WidgetSize, Dashboard } from '@pnp/spfx-controls-react/lib/Dashboard';
import IMemberListItem from '../ClaringtonInterfaces/IMemberListItem';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GetMember, GetMembers, OnFormatDate } from '../HelperMethods/MyHelperMethods';
import { CommitteeMemberContactDetails, CommitteeMemberTermHistory } from './MemberDetailsComponent';


export interface ICommitteeMemberDashboardProps {
    description?: string;
    memberId?: number;
    context: WebPartContext;
}

export interface ICommitteeMemberDashboardState {
    members: IMemberListItem[];
    selectedMember?: IMemberListItem;
}

export class CommitteeMemberDashboard extends React.Component<ICommitteeMemberDashboardProps, ICommitteeMemberDashboardState> {

    constructor(props: ICommitteeMemberDashboardProps) {
        super(props);
        this.state = {
            members: undefined,
        };

        GetMembers().then(members => {
            this.setState({
                members: members
            });
        }).catch(reason => {
            console.error('Failed to get members');
            console.error(reason);
            this.setState({ members: undefined });
        });

        if (this.props.memberId) {
            GetMember(this.props.memberId).then(value => {
                this.setState({ selectedMember: value });
            }).catch(reason => {
                console.error('Failed to get selected member');
                console.error(reason);
                this.setState({ selectedMember: undefined });
            });
        }
    }

    public render(): React.ReactElement<ICommitteeMemberDashboardProps> {
        const linkExample = { href: "#" };
        const calloutItemsExample = [
            {
                id: "action_1",
                title: "Info",
                icon: <Icon iconName={'Edit'} />,
            },
            { id: "action_2", title: "Popup", icon: <Icon iconName={'Add'} /> },
        ];

        return <div>
            {
                this.state.members &&
                <ComboBox
                    label={'Select Member'}
                    options={this.state.members.map((member: IMemberListItem) => {
                        return { key: member.ID, text: member.Title, data: { ...member } };
                    })}
                    onChange={(event, option) => {
                        this.setState({ selectedMember: undefined });
                        GetMember(Number(option.key))
                            .then(member => {
                                this.setState({ selectedMember: member });
                            }).catch(reason => {
                                console.error('Failed to get member!');
                                console.error(reason);
                            });
                    }}
                    defaultSelectedKey={this.props.memberId ? this.props.memberId : undefined}
                />
            }
            {
                this.state.selectedMember &&
                <Dashboard
                    widgets={[{
                        title: this.state.selectedMember.Title,
                        desc: `Last updated ${OnFormatDate(new Date(this.state.selectedMember.Modified))}`,
                        widgetActionGroup: calloutItemsExample,
                        size: WidgetSize.Triple,
                        body: [
                            {
                                id: "t1",
                                title: "Tab 1",
                                content: (
                                    <CommitteeMemberContactDetails member={this.state.selectedMember} />
                                ),
                            }
                        ],
                        link: linkExample,
                    },
                    {
                        title: "Committee History",
                        size: WidgetSize.Single,
                        body: [{
                            id: 'id',
                            title: 'Committee History',
                            content: (<CommitteeMemberTermHistory memberID={this.state.selectedMember.ID} context={this.props.context} />)
                        }],
                        link: linkExample,
                    },
                    {
                        title: "Card 3",
                        size: WidgetSize.Double,
                        link: linkExample,
                    }]} />
            }
        </div>;
    }
}

