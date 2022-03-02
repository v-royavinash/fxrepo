import { Button, ChevronStartIcon, Flex, FormDropdown, FormInput, FormTextArea, Loader } from "@fluentui/react-northstar";
import { LocalizationHelper, PeoplePicker } from '@microsoft/mgt-react';
import { Client } from "@microsoft/microsoft-graph-client";
import 'bootstrap/dist/css/bootstrap.min.css';
import moment from "moment";
import * as React from "react";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import { v4 as uuidv4 } from "uuid";
import CommonService from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import '../scss/IncidentDetails.module.scss';
import {
    ChannelCreationResult, ChannelCreationStatus, CreateIncidentEntity,
    IInputValidationStates, ITeamChannel, ITeamCreatedResponse,
    RoleAssignments
} from "./ICreateIncident";
import { IInputRegexValidationStates } from '../common/CommonService';
import { ITooltipHostStyles, TooltipHost } from "@fluentui/react/lib/Tooltip";
import { Icon } from "@fluentui/react/lib/Icon";

const calloutProps = { gapSpace: 0 };

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', cursor: 'pointer' } };

export interface IIncidentDetailsProps {
    graph: Client;
    tenantName: string;
    siteId: string;
    onBackClick(showMessageBar: boolean): void;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    localeStrings: any;
    currentUserId: string;
}

export interface IIncidentDetailsState {
    dropdownOptions: any;
    createIncDetailsItem: CreateIncidentEntity;
    newRoleString: string;
    roleAssignments: RoleAssignments[];
    showLoader: boolean;
    loaderMessage: string;
    inputValidation: IInputValidationStates;
    inputRegexValidation: IInputRegexValidationStates;
    isCreateNewRoleBtnDisabled: boolean;
    isAddRoleAssignmentBtnDisabled: boolean;
    isDesktop: boolean;
    formOpacity: number;
    eocAppId: string;
    selectedUsers: any
}

// sets the initial values for required fields validation object
const getInputValildationInitialState = (): IInputValidationStates => {
    return {
        incidentNameHasError: false,
        incidentStatusHasError: false,
        incidentLocationHasError: false,
        incidentTypeHasError: false,
        incidentDescriptionHasError: false,
        incidentStartDateTimeHasError: false,
        incidentCommandarHasError: false,
    };
};

class IncidentDetails extends React.PureComponent<IIncidentDetailsProps, IIncidentDetailsState> {
    constructor(props: IIncidentDetailsProps) {
        super(props);
        this.state = {
            dropdownOptions: '',
            createIncDetailsItem: new CreateIncidentEntity(),
            newRoleString: '',
            roleAssignments: [],
            showLoader: true,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            inputValidation: getInputValildationInitialState(),
            inputRegexValidation: this.dataService.getInputRegexValildationInitialState(),
            isCreateNewRoleBtnDisabled: true,
            isAddRoleAssignmentBtnDisabled: true,
            isDesktop: true,
            formOpacity: 0.5,
            eocAppId: "",
            selectedUsers: []
        };
        this.onRoleChange = this.onRoleChange.bind(this);
        this.onTextInputChange = this.onTextInputChange.bind(this);
        this.handleIncCommanderChange = this.handleIncCommanderChange.bind(this);
        this.onAddNewRoleChange = this.onAddNewRoleChange.bind(this);
        this.onIncidentTypeChange = this.onIncidentTypeChange.bind(this);
        this.onIncidentStatusChange = this.onIncidentStatusChange.bind(this);
        this.onRoleChange = this.onRoleChange.bind(this);


        // localized messages for people pickers
        LocalizationHelper.strings = {
            _components: {
                'people-picker': {
                    noResultsFound: this.props.localeStrings.peoplePickerNoResult,
                    loadingMessage: this.props.localeStrings.peoplePickerLoader
                }
            }
        }
    }

    private dataService = new CommonService();
    private graphEndpoint = "";

    public async componentDidMount() {
        this.getDropdownOptions();
        //Event listener for screen resizing
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
    }

    //Function for screen Resizing
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth })

    componentWillUnmount() {

        //Event listener for screen resizing
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // get dropdown options 
    getDropdownOptions = async () => {
        try {
            const incStatusGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incStatusList}/items?$expand=fields&$Top=5000`;
            const incTypeGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incTypeList}/items?$expand=fields&$Top=5000`;
            const roleGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleAssignmentList}/items?$expand=fields&$Top=5000`;

            const statusOptionsPromise = this.dataService.getDropdownOptions(incStatusGraphEndpoint, this.props.graph);
            const typeOptionsPromise = this.dataService.getDropdownOptions(incTypeGraphEndpoint, this.props.graph);
            const roleOptionsPromise = this.dataService.getDropdownOptions(roleGraphEndpoint, this.props.graph);

            await Promise.all([statusOptionsPromise, typeOptionsPromise, roleOptionsPromise])
                .then(([statusOptions, typeOptions, roleOptions]) => {
                    const optionsArr: any = [];
                    // remove "Closed" status from options
                    optionsArr.statusOptions = statusOptions.filter((status: string) => status !== constants.closed);
                    optionsArr.typeOptions = typeOptions.sort();
                    optionsArr.roleOptions = roleOptions.sort();

                    let incInfo = { ...this.state.createIncDetailsItem };
                    let inputValidationObj = this.state.inputValidation;
                    if (incInfo) {
                        if (incInfo) {
                            incInfo["incidentStatus"] = constants.active;
                            inputValidationObj.incidentStatusHasError = false;
                        }
                    }

                    this.setState({
                        dropdownOptions: optionsArr,
                        showLoader: false,
                        createIncDetailsItem: incInfo,
                        inputValidation: inputValidationObj,
                        formOpacity: 1
                    })
                }, (error: any): void => {
                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.dropdownRetrievalFailedErrMsg, constants.messageBarType.error);
                    this.setState({
                        showLoader: false,
                        formOpacity: 1
                    })
                });
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_GetDropdownOptions \n",
                JSON.stringify(error)
            );
        }
    }

    // on incident commander change
    private handleIncCommanderChange = (selectedValue: any) => {
        let incInfo = { ...this.state.createIncDetailsItem };
        if (incInfo) {
            let inputValidationObj = this.state.inputValidation;
            if (selectedValue.detail.length > 0) {
                inputValidationObj.incidentCommandarHasError = false;
            }
            else {
                inputValidationObj.incidentCommandarHasError = true;
            }
            // create user object for incident commander
            incInfo.incidentCommander = {
                userName: selectedValue.detail[0] ? selectedValue.detail[0].displayName : '',
                userEmail: selectedValue.detail[0] ? selectedValue.detail[0].userPrincipalName : '',
                userId: selectedValue.detail[0] ? selectedValue.detail[0].id : ''
            }
            this.setState({ createIncDetailsItem: incInfo, inputValidation: inputValidationObj });
        }
    };

    // on change handler for text input changes
    private onTextInputChange = (event: any, key: string) => {
        let incInfo = { ...this.state.createIncDetailsItem };
        let inputValidationObj = this.state.inputValidation;
        let regexValidationObj = this.state.inputRegexValidation;
        if (incInfo) {
            switch (key) {
                case "incidentName":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentNameHasError = false;
                    }
                    else {
                        inputValidationObj.incidentNameHasError = true;
                    }
                    this.setState({
                        createIncDetailsItem: incInfo,
                        inputValidation: inputValidationObj
                    })
                    break;
                case "startDateTime":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentStartDateTimeHasError = false;
                    }
                    else {
                        inputValidationObj.incidentStartDateTimeHasError = true;
                    }
                    this.setState({ createIncDetailsItem: incInfo, inputValidation: inputValidationObj })
                    break;
                case "location":
                    incInfo[key] = event.target.value;

                    // check for required field validation
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentLocationHasError = false;
                    }
                    else {
                        inputValidationObj.incidentLocationHasError = true;
                    }
                    this.setState({
                        createIncDetailsItem: incInfo,
                        inputValidation: inputValidationObj
                    })
                    break;
                case "assignedUser":
                    incInfo[key] = event.target.value;
                    this.setState({ createIncDetailsItem: incInfo });
                    break;
                case "incidentDesc":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentDescriptionHasError = false;
                    }
                    else {
                        inputValidationObj.incidentDescriptionHasError = true;
                    }
                    this.setState({ createIncDetailsItem: incInfo, inputValidation: inputValidationObj })
                    break;

                default:
                    break;
            }
        }
    }

    // update state for new role string
    private onAddNewRoleChange = (event: any) => {
        let isButtonDisabled = true;
        if (event.target.value && event.target.value.length > 0) {
            isButtonDisabled = false;
        }
        this.setState({ newRoleString: event.target.value, isCreateNewRoleBtnDisabled: isButtonDisabled });
    }

    // on incident type dropdown value change
    private onIncidentTypeChange = (event: any, selectedValue: any) => {
        let incInfo = this.state.createIncDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.createIncDetailsItem };
            if (incInfo) {
                incInfo["incidentType"] = selectedValue.value;
                let inputValidationObj = this.state.inputValidation;
                inputValidationObj.incidentTypeHasError = false;
                this.setState({ createIncDetailsItem: incInfo, inputValidation: inputValidationObj })
            }
        }
    }

    // on incident status dropdown value change
    private onIncidentStatusChange = (event: any, selectedValue: any) => {
        let incInfo = this.state.createIncDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.createIncDetailsItem };
            if (incInfo) {
                incInfo["incidentStatus"] = selectedValue.value;
                let inputValidationObj = this.state.inputValidation;
                inputValidationObj.incidentStatusHasError = false;
                this.setState({ createIncDetailsItem: incInfo, inputValidation: inputValidationObj })
            }
        }
    }

    // on role dropdown value change
    private onRoleChange = (event: any, selectedRole: any) => {
        let incInfo = this.state.createIncDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.createIncDetailsItem };
            if (incInfo) {
                incInfo["selectedRole"] = selectedRole.value;
                this.setState({ createIncDetailsItem: incInfo }, (() => this.checkAddRoleBtnState()))
            }
        }
    }

    // connect with service to create new role in Role Assignments list
    private addNewRole = async () => {

        this.setState({
            showLoader: true,
            formOpacity: 0.5
        })
        // create graph endpoint for role assignment list
        this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleAssignmentList}/items`;

        // create new item object to add the role
        const listItem = {
            fields: {
                Title: this.state.newRoleString
            }
        };

        try {
            const addedRole = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, listItem);

            if (addedRole) {
                const arr: any = { ...this.state.dropdownOptions };
                arr["roleOptions"].push(this.state.newRoleString);

                let incInfo = this.state.createIncDetailsItem;
                if (incInfo) {
                    let incInfo = { ...this.state.createIncDetailsItem };
                    if (incInfo) {
                        incInfo["selectedRole"] = this.state.newRoleString;
                        this.setState({
                            createIncDetailsItem: incInfo, newRoleString: "", showLoader: false,
                            formOpacity: 1
                        })
                        this.props.showMessageBar(this.props.localeStrings.addRoleSuccessMessage, "success");
                    }
                    else {
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        })
                    }
                }
            }
            else {
                this.setState({
                    showLoader: false,
                    formOpacity: 1
                })
            }
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_AddNewRole \n",
                JSON.stringify(error)
            );
        }
    }

    // on assigned user change
    private handleAssignedUserChange = (selectedValue: any) => {
        let incInfo = { ...this.state.createIncDetailsItem };
        const selectedUsersArr: any = [];
        if (incInfo) {
            incInfo["assignedUser"] = selectedValue.detail.map((user: any) => {
                selectedUsersArr.push({
                    displayName: user.displayName,
                    userPrincipalName: user.userPrincipalName,
                    id: user.id
                });
                return {
                    "userName": user ? user.displayName : "",
                    "userEmail": user ? user.userPrincipalName : "",
                    "userId": user ? user.id : "",
                }
            });

            this.setState({ createIncDetailsItem: incInfo, selectedUsers: selectedUsersArr });
            this.checkAddRoleBtnState();
        }
    };

    // update the role assignment array
    private addRoleAssignment = () => {
        let roleAssignment = [...this.state.roleAssignments];
        let userDetailsObj: any = [];
        let userNameString = "";
        // push roles into array to create role object
        this.state.createIncDetailsItem.assignedUser.forEach(assignedUser => {
            userNameString += assignedUser.userName + ", ";
            userDetailsObj.push({
                userName: assignedUser.userName,
                userEmail: assignedUser.userEmail,
                userId: assignedUser.userId,
            });
        });
        userNameString = userNameString.trim();
        userNameString = userNameString.slice(0, -1);

        roleAssignment.push({
            role: this.state.createIncDetailsItem.selectedRole,
            userNamesString: userNameString,
            userDetailsObj: userDetailsObj
        })

        // clear roles control values
        let incInfo = this.state.createIncDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.createIncDetailsItem };
            if (incInfo) {
                incInfo["selectedRole"] = "";
                this.setState({
                    roleAssignments: roleAssignment, createIncDetailsItem: incInfo,
                    selectedUsers: [],
                    isAddRoleAssignmentBtnDisabled: true
                })
            }
        }
    }

    // change add role assignment button disable state
    private checkAddRoleBtnState = () => {
        if (this.state.createIncDetailsItem.selectedRole !== "" &&
            this.state.selectedUsers && this.state.selectedUsers.length > 0) {
            this.setState({
                isAddRoleAssignmentBtnDisabled: false
            })
        }
        else {
            this.setState({
                isAddRoleAssignmentBtnDisabled: true
            })
        }
    }

    // delete added role from RoleAssignment object
    private deleteRoleItem = (itemIndex: number) => {
        let assignments = [...this.state.roleAssignments];
        assignments.splice(itemIndex, 1);
        this.setState({ roleAssignments: assignments });
    }

    // create new entry in incident transaction list
    private createNewIncident = async () => {
        this.scrollToTop();
        // incident info object
        let incidentInfo: CreateIncidentEntity = this.state.createIncDetailsItem;
        this.props.hideMessageBar();
        this.setState({
            showLoader: true,
            formOpacity: 0.5,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            inputRegexValidation: this.dataService.getInputRegexValildationInitialState(),
            inputValidation: getInputValildationInitialState()
        })

        // validate for required fields
        if (!this.requiredFieldValidation(incidentInfo)) {
            this.props.showMessageBar(this.props.localeStrings.reqFieldErrorMessage, constants.messageBarType.error);
        }
        else {
            try {
                // validate input strings for incident name and location
                const regexValidation = this.dataService.regexValidation(incidentInfo);
                if (regexValidation.incidentLocationHasError || regexValidation.incidentNameHasError) {
                    this.props.showMessageBar(this.props.localeStrings.regexErrorMessage, constants.messageBarType.error);
                    this.setState({
                        inputRegexValidation: regexValidation,
                        showLoader: false,
                        formOpacity: 1
                    });
                }
                else {
                    let incNameStr = incidentInfo.incidentName.replace(/'/g, "''");
                    // create graph endpoint for querying Incident Transaction list
                    this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentsList}/items?$expand=fields&$filter=fields/IncidentName eq '${incNameStr}'`;

                    // show error if incident with same name already exists
                    if (await this.dataService.getIncident(this.graphEndpoint, this.props.graph)) {
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        });
                        this.props.showMessageBar(this.props.localeStrings.duplicateIncidentName, constants.messageBarType.error);
                    }
                    else {
                        try {
                            this.setState({
                                loaderMessage: this.props.localeStrings.incidentCreationLoaderMessage
                            });
                            // prepare the role assignment object which will be stored in 
                            // incident transaction list in string format
                            let roleAssignment = "";
                            this.state.roleAssignments.forEach(roles => {
                                roleAssignment += roles.role + " - " + roles.userNamesString + "; ";
                            });

                            // create object to be passed in graph query
                            const incidentInfoObj: any = {
                                fields: {
                                    Title: incidentInfo.incidentName,
                                    Description: incidentInfo.incidentDesc,
                                    IncidentType: incidentInfo.incidentType,
                                    IncidentStatus: incidentInfo.incidentStatus,
                                    TeamId: "",
                                    StartDateTime: incidentInfo.startDateTime + ":00Z",
                                    Location: incidentInfo.location,
                                    IncidentName: incidentInfo.incidentName,
                                    RoleAssignment: roleAssignment.trim(),
                                    IncidentCommander: incidentInfo.incidentCommander.userName
                                }
                            }

                            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentsList}/items`;

                            const incidentAdded = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, incidentInfoObj);

                            // check if incident is created
                            if (incidentAdded) {
                                console.log(constants.infoLogPrefix + "Incident Created");
                                try {
                                    // call method to update the incident id with custom value
                                    const incUpdated = await this.updatedIncidentId(incidentAdded.id);

                                    if (incUpdated) {
                                        console.log(constants.infoLogPrefix + "Incident Id Updated");
                                        // call the wrapper method to perform Teams related operations
                                        await this.createTeamAndChannels(incUpdated.IncidentId, incidentAdded.id);
                                    }
                                    else {
                                        // delete the incident if incident id updation fails
                                        await this.deleteIncident(incidentAdded.id);
                                        this.setState({
                                            showLoader: false,
                                            formOpacity: 1
                                        });
                                        this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                                    }
                                } catch (error) {
                                    console.error(
                                        constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                                        JSON.stringify(error)
                                    );
                                    // delete the item if error occured
                                    await this.deleteIncident(incidentAdded.id);
                                    this.setState({
                                        showLoader: false,
                                        formOpacity: 1
                                    });
                                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                                }
                            }
                            else {
                                this.setState({
                                    showLoader: false,
                                    formOpacity: 1
                                });
                                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                            }
                        } catch (error) {
                            console.error(
                                constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                                JSON.stringify(error)
                            );
                            this.setState({
                                showLoader: false,
                                formOpacity: 1
                            });
                            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                        }
                    }
                }
            } catch (error) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                    JSON.stringify(error)
                );
            }
        }
    }

    // perform required fields validation
    private requiredFieldValidation = (incidentInfo: CreateIncidentEntity): boolean => {
        let inputValidationObj = getInputValildationInitialState();
        let reqFieldValidationSuccess = true;
        if (incidentInfo.incidentName === "" || incidentInfo.incidentName === undefined) {
            inputValidationObj.incidentNameHasError = true;
        }
        if (incidentInfo.incidentType === "" || incidentInfo.incidentType === undefined) {
            inputValidationObj.incidentTypeHasError = true;
        }
        if (incidentInfo.startDateTime === "" || incidentInfo.startDateTime === undefined) {
            inputValidationObj.incidentStartDateTimeHasError = true;
        }
        if (incidentInfo.incidentStatus === "" || incidentInfo.incidentStatus === undefined) {
            inputValidationObj.incidentStatusHasError = true;
        }
        if (incidentInfo.location === "" || incidentInfo.location === undefined) {
            inputValidationObj.incidentLocationHasError = true;
        }
        if (incidentInfo.incidentDesc === "" || incidentInfo.incidentDesc === undefined) {
            inputValidationObj.incidentDescriptionHasError = true;
        }
        if ((incidentInfo.incidentCommander === undefined)) {
            inputValidationObj.incidentCommandarHasError = true;
        }

        if (inputValidationObj.incidentNameHasError || inputValidationObj.incidentTypeHasError ||
            inputValidationObj.incidentStartDateTimeHasError || inputValidationObj.incidentStatusHasError ||
            inputValidationObj.incidentLocationHasError || inputValidationObj.incidentDescriptionHasError ||
            inputValidationObj.incidentCommandarHasError) {
            this.setState({
                inputValidation: inputValidationObj,
                showLoader: false,
                formOpacity: 1
            });
            reqFieldValidationSuccess = false;
        }
        return reqFieldValidationSuccess;
    }

    // method to delay the operation by adding timeout
    private timeout = (delay: number): Promise<any> => {
        return new Promise(res => setTimeout(res, delay));
    }

    // wrapper method to perform teams related operations
    private async createTeamAndChannels(incidentId: any, listItemId: number): Promise<any> {
        return new Promise(async (resolve, reject) => {
            // response object for Teams creation
            let teamCreationResult: ITeamCreatedResponse = this.getITeamCreatedResponseDefaultValue();

            console.log(constants.infoLogPrefix + "Teams group creation start");
            // call method to create Teams group
            this.createTeamGroup(incidentId).then(async (groupInfo) => {
                try {
                    console.log(constants.infoLogPrefix + "Teams group created on - " + new Date());

                    // wait for 2 seconds to ensure team group is available via graph API
                    await this.timeout(2000);

                    // create associated team with the group
                    const teamInfo = await this.createTeam(groupInfo);
                    if (teamInfo.status) {
                        console.log(constants.infoLogPrefix + "Teams created on - " + new Date());
                        // create channels
                        const channelCreatedInfo: any = await this.createChannels(teamInfo.data);
                        console.log(constants.infoLogPrefix + "channels created");

                        const siteURL = "https://" + this.props.tenantName + ".sharepoint.com/sites/" + groupInfo.mailNickname;

                        // create assessment channel and tab
                        await this.createAssessmentChannelAndTab(groupInfo.id, siteURL, groupInfo.mailNickname);

                        console.log(constants.infoLogPrefix + "Assessment Channel and tab created");

                        const siteBaseURL = "https://" + this.props.tenantName + ".sharepoint.com/sites/";

                        // create news channel and tab
                        await this.createNewsTab(groupInfo, siteBaseURL);
                        console.log(constants.infoLogPrefix + "News tab created");

                        // create URL to get site Id
                        const urlForSiteId = graphConfig.spSiteGraphEndpoint + this.props.tenantName + ".sharepoint.com:/sites/" + groupInfo.mailNickname + "?$select=id";

                        const siteDetails = await this.dataService.getGraphData(urlForSiteId, this.props.graph);
                        console.log(constants.infoLogPrefix + "Site details retrieved");

                        // call method to create assessment list
                        const assessmentList = await this.createAssessmentList(groupInfo.mailNickname, siteDetails.id);
                        console.log(constants.infoLogPrefix + "Assessment list created");

                        // get all columns to get status column ID
                        const allColumnsGraphEndpoint = graphConfig.sitesGraphEndpoint + "/" + siteDetails.id + graphConfig.listsGraphEndpoint + "/" + assessmentList.id + graphConfig.columnsGraphEndpoint;

                        const allColumnsResponse = await this.dataService.getGraphData(allColumnsGraphEndpoint, this.props.graph);
                        console.log(constants.infoLogPrefix + "All columns retrieved");

                        // check if object is having values
                        if (allColumnsResponse && allColumnsResponse.value.length > 0) {
                            // filter to get status column
                            const statusColumn = allColumnsResponse.value.filter((column: any) => {
                                return column.name === "Status"
                            });

                            const statusColGraphEndpoint = allColumnsGraphEndpoint + "/" + statusColumn[0].id;
                            // apply formatting to status column
                            await this.dataService.sendGraphPatchRequest(statusColGraphEndpoint, this.props.graph, { CustomFormatter: siteConfig.AssessmentListStatusFormat });
                            console.log(constants.infoLogPrefix + "Column formatting success");
                        }

                        const updateItemObj = {
                            TeamId: teamInfo.id,
                            TeamWebURL: teamInfo.data.webUrl
                        }

                        await this.updatedTeamIdInList(listItemId, updateItemObj);
                        console.log(constants.infoLogPrefix + "List item updated with Team Id");

                        // create the tags for incident commander and each selected roles
                        await this.createTagObject(teamInfo);

                        // update the results object
                        teamCreationResult.fullyDone = (channelCreatedInfo.is_fully_created ? true : false);
                        teamCreationResult.partiallyDone = !(channelCreatedInfo.is_fully_created);
                        teamCreationResult.error.channelCreations = channelCreatedInfo;
                        teamCreationResult.teamInfo = groupInfo;

                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        })
                        this.props.showMessageBar(this.props.localeStrings.incidentCreationSuccessMessage, constants.messageBarType.success);
                        this.props.onBackClick(true);
                    }
                    else {
                        // delete the group if some error occured
                        await this.deleteTeamGroup(groupInfo.id);
                        // delete the item if error occured
                        await this.deleteIncident(listItemId);

                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        })
                        this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                    }
                } catch (error) {
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_createTeamAndChannels \n",
                        JSON.stringify(error)
                    );
                    // delete the group if some error occured
                    await this.deleteTeamGroup(groupInfo.id);
                    // delete the item if error occured
                    await this.deleteIncident(listItemId);

                    this.setState({
                        showLoader: false,
                        formOpacity: 1
                    })
                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                }
            }).catch((error) => {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_createTeamAndChannels \n",
                    JSON.stringify(error)
                );
                // delete the item if error occured
                this.deleteIncident(listItemId);

                this.setState({
                    showLoader: false,
                    formOpacity: 1
                });
                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
            });
        })
    }

    // Initialize the Teams creation response object
    private getITeamCreatedResponseDefaultValue(): ITeamCreatedResponse {
        let _result: ITeamCreatedResponse = {
            fullyDone: false,
            partiallyDone: false,
            allFailed: false,
            teamInfo: "",
            error: {
                channelCreations: [],
                appInstallation: [],
                memberCreations: [],
                allFail: []
            }
        };
        return _result;
    }

    // updates incident ID based on created item Id
    private updatedIncidentId = async (itemId: number): Promise<any> => {
        try {
            const updateValues = {
                IncidentId: itemId
            }
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${itemId}/fields`;

            const updatedIncident = await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, updateValues);
            return updatedIncident;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_UpdatedIncidentId \n",
                JSON.stringify(error)
            );
        }
    }

    // updates incident ID based on created item Id
    private updatedTeamIdInList = async (itemId: number, updateItemObj: any): Promise<any> => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${itemId}/fields`;
            const updatedIncident = await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, updateItemObj);
            return updatedIncident;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_UpdatedTeamIdInList \n",
                JSON.stringify(error)
            );
        }
    }

    // create a Teams group
    private createTeamGroup = async (incId: string): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            try {
                let incDetails = this.state.createIncDetailsItem;
                // update the date format
                incDetails.startDateTime = moment(this.state.createIncDetailsItem.startDateTime).format("DDMMMYYYY");

                // create members array
                const membersArr: any = [];
                this.state.roleAssignments.forEach(roles => {
                    roles.userDetailsObj.forEach(user => {
                        if (membersArr.indexOf(graphConfig.usersGraphEndpoint + user.userId) === -1) {
                            membersArr.push(graphConfig.usersGraphEndpoint + user.userId);
                        }
                    });
                });

                const ownerArr: any = [];
                ownerArr.push(graphConfig.usersGraphEndpoint + incDetails.incidentCommander.userId);

                // add current user as a owner if already not present so that we can perform teams creation
                // and sharepoint site related operations on associated team site
                if (ownerArr.indexOf(graphConfig.usersGraphEndpoint + this.props.currentUserId) === -1) {
                    ownerArr.push(graphConfig.usersGraphEndpoint + this.props.currentUserId)
                }

                if (membersArr.length > 0) {
                    // create object to create teams group
                    let incidentobj = {
                        displayName: `${constants.teamEOCPrefix}-${incId}-${incDetails.incidentType}-${incDetails.startDateTime}`,
                        mailNickname: `${constants.teamEOCPrefix}_${incId}`,
                        description: incDetails.incidentDesc,
                        visibility: "Private",
                        groupTypes: ["Unified"],
                        mailEnabled: true,
                        securityEnabled: true,
                        "members@odata.bind": membersArr,
                        "owners@odata.bind": ownerArr
                    }
                    // call method to create team group
                    let groupResponse = await this.dataService.sendGraphPostRequest(graphConfig.teamGroupsGraphEndpoint, this.props.graph, incidentobj);
                    resolve(groupResponse);
                }
                else {
                    // create object to create teams group
                    let incidentobj = {
                        displayName: `${constants.teamEOCPrefix}-${incId}-${incDetails.incidentType}-${incDetails.startDateTime}`,
                        mailNickname: `${constants.teamEOCPrefix}_${incId}`,
                        description: incDetails.incidentDesc,
                        visibility: "Private",
                        groupTypes: ["Unified"],
                        mailEnabled: true,
                        securityEnabled: true,
                        "owners@odata.bind": ownerArr
                    }
                    // call method to create team group
                    let groupResponse = await this.dataService.sendGraphPostRequest(graphConfig.teamGroupsGraphEndpoint, this.props.graph, incidentobj);
                    resolve(groupResponse);
                }
            }
            catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateTeamGroup \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                console.error("EOC App - CreateTeamGroup_Failed to create teams group \n" + ex);
            }

        })
    }

    // create Team associated with Teams group
    private async createTeam(groupInfo: any): Promise<any> {
        return new Promise(async (resolve) => {
            let maxTeamCreationAttempt = 5, isTeamCreated = false;

            let result = {
                status: false,
                data: {}
            };

            // loop till the team is created
            // attempting multiple times as sometimes teams group doesn't reflect immediately after creation
            while (isTeamCreated === false && maxTeamCreationAttempt > 0) {
                // let dataService = new CommonService();
                try {
                    // create the team setting object
                    let teamSettings = JSON.stringify(this.getTeamSettings());
                    this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + groupInfo.id + graphConfig.teamGraphEndpoint

                    // call method to create team
                    let updatedTeamInfo = await this.dataService.sendGraphPutRequest(this.graphEndpoint, this.props.graph, teamSettings)

                    // update the result object
                    if (updatedTeamInfo) {
                        console.log(constants.infoLogPrefix + "Teams created on - " + new Date());
                        isTeamCreated = true;
                        result.data = updatedTeamInfo;
                        result.status = true;
                    }
                } catch (updationError: any) {
                    console.log(constants.infoLogPrefix + "Teams creation failed on - " + new Date());
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_CreateTeam \n",
                        JSON.stringify(updationError)
                    );
                    if (updationError.statusCode === 409 && updationError.message === "Team already exists") {
                        isTeamCreated = true;
                        this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + groupInfo.id;
                        result.data = await this.dataService.getGraphData(this.graphEndpoint, this.props.graph)
                    }
                }
                maxTeamCreationAttempt--;
                await this.timeout(10000);
            }
            console.log(constants.infoLogPrefix + "createTeam_No Of Attempt", (5 - maxTeamCreationAttempt), result);
            resolve(result);
        });
    }

    // set the Teams properties
    private getTeamSettings = (): any => {
        return {
            "memberSettings": {
                "allowCreateUpdateChannels": true,
                "allowDeleteChannels": true,
                "allowAddRemoveApps": true,
                "allowCreateUpdateRemoveTabs": true,
                "allowCreateUpdateRemoveConnectors": true
            },
            "guestSettings": {
                "allowCreateUpdateChannels": true,
                "allowDeleteChannels": true
            },
            "messagingSettings": {
                "allowUserEditMessages": true,
                "allowUserDeleteMessages": true,
                "allowOwnerDeleteMessages": true,
                "allowTeamMentions": true,
                "allowChannelMentions": true
            },
            "funSettings": {
                "allowGiphy": true,
                "giphyContentRating": "strict",
                "allowStickersAndMemes": true,
                "allowCustomMemes": true
            }
        };
    }

    // get channels to be created
    private getFixedChannel(): Array<ITeamChannel> {
        let res: Array<ITeamChannel> = [];
        res.push({
            "displayName": "Logistics",
        });
        res.push({
            "displayName": "Planning",
        });
        res.push({
            "displayName": "Recovery",
        });
        res.push({
            "displayName": "Urgent",
        });
        return res;
    }

    // create channels
    private async createChannels(group_details: any): Promise<any> {
        //some time graph api does't create the channel 
        //thats why we need to re-try 2 time if again it failed then need to take this into failed item. otherwise simply add into 
        //created list, we need to show end-use if something failed then need to pop those error.

        let channels = this.getFixedChannel();
        let result: ChannelCreationResult = {
            isFullyCreated: false,
            isPartiallyCreated: false,
            failedEntries: [],
            successEntries: []
        };
        const MAX_NUMBER_OF_ATTEMPT = 3;
        let noOfAttempt = 1;
        return new Promise(async (resolve, reject) => {
            let allDone = false;
            let counter = 0;

            // loop atlease 3 times or till the channel is created
            while (!allDone) {
                let channel = channels[counter];
                try {
                    // const dataService = new CommonService();
                    this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + group_details.id + graphConfig.channelsGraphEndpoint;
                    let createdChannel = await this.dataService.createChannel(this.graphEndpoint, this.props.graph, channel)

                    if (createdChannel) {
                        // set channel object
                        let channelObj: ChannelCreationStatus = {
                            channelName: channel.displayName,
                            isCreated: true,
                            noOfCreationAttempt: noOfAttempt,
                            rawData: createdChannel
                        };
                        noOfAttempt = 1;
                        result.successEntries.push(channelObj);
                    }
                    counter++;
                } catch (ex: any) {
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_CreateChannels \n",
                        JSON.stringify(ex)
                    );
                    if (noOfAttempt >= MAX_NUMBER_OF_ATTEMPT) {
                        let channelObj: ChannelCreationStatus = {
                            channelName: channel.displayName,
                            isCreated: false,
                            noOfCreationAttempt: noOfAttempt,
                            rawData: ex.message
                        };
                        noOfAttempt = 1;
                        result.isFullyCreated = false;
                        result.failedEntries.push(channelObj);
                        counter++;
                    } else {
                        noOfAttempt++;
                    }
                }
                allDone = (channels.length - 1) === counter;
            }
            result.isFullyCreated = result.failedEntries.length === 0 ? true : false;
            resolve(result);
        });
    }

    // create assessment channel and tab
    private async createAssessmentChannelAndTab(team_id: string, site_base_url: string, site_name: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            try {
                const channelGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_id + graphConfig.channelsGraphEndpoint;
                const channelObj = {
                    "displayName": constants.Assessment,
                    isFavoriteByDefault: true
                };

                const channelResult = await this.dataService.createChannel(channelGraphEndpoint, this.props.graph, channelObj);
                console.log(constants.infoLogPrefix + "Assessment channel created");

                const tabGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_id + graphConfig.channelsGraphEndpoint + "/" + channelResult.id + graphConfig.tabsGraphEndpoint;

                //Associate Assessment via sharepoint app
                const assessmentTabObj = {
                    "displayName": constants.GroundAssessments,
                    "teamsApp@odata.bind": graphConfig.assessmentTabTeamsAppIdGraphEndpoint,
                    "configuration": {
                        "entityId": uuidv4(),
                        "contentUrl": `${site_base_url}/_layouts/15/listallitems.aspx?listUrl=/sites/${site_name}/Lists/Assessments&app=teamslist&v=2`,
                        "removeUrl": null,
                        "websiteUrl": null
                    }
                }

                await this.dataService.sendGraphPostRequest(tabGraphEndpoint, this.props.graph, assessmentTabObj);
                console.log(constants.infoLogPrefix + "list view added to assessment tab");
                resolve({
                    status: true,
                    message: "channel and tab created also installed app into tab"
                });
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateAssessmentChannelAndTab \n",
                    JSON.stringify(ex)
                );
                reject(ex);

            }
        });
    }

    // create News tab
    private createNewsTab(team_info: any, siteBaseURL: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            try {
                this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_info.id + graphConfig.channelsGraphEndpoint;

                const tabObj = {
                    "displayName": constants.Announcements,
                    "description": "",
                    isFavoriteByDefault: true
                };
                const channelResult = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, tabObj);
                console.log(constants.infoLogPrefix + "News tab created");
                // get the app ID
                //const app = await this.getTeamEOCApp();

                const addTabObj = {
                    "displayName": constants.News,
                    "teamsApp@odata.bind": graphConfig.newsTabTeamsAppIdGraphEndpoint,
                    "configuration": {
                        "entityId": uuidv4(),
                        "contentUrl": `${siteBaseURL}${team_info.mailNickname}/_layouts/15/news.aspx`,
                        "removeUrl": null,
                        "websiteUrl": `${siteBaseURL}${team_info.mailNickname}/_layouts/15/news.aspx`
                    }
                }
                const addTabGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_info.id + graphConfig.channelsGraphEndpoint + "/" + channelResult.id + graphConfig.tabsGraphEndpoint;

                // calling a generic method which is send a post query to the graph endpoint
                await this.dataService.sendGraphPostRequest(addTabGraphEndpoint, this.props.graph, addTabObj);
                console.log(constants.infoLogPrefix + "News page added to news tab");
                resolve({
                    status: true,
                    message: "channel and tab created also installed app into tab"
                });
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateNewsTab \n",
                    JSON.stringify(ex)
                );
                reject(ex);
            }
        });
    }

    // this method creates assessment list in the new team site for incident
    private async createAssessmentList(siteName: string, siteId: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            /* List,Field,View Creation */
            try {
                let listColumns: any = [];

                siteConfig.lists[0].columns.forEach(column => {
                    listColumns.push(column);
                });

                let listSchema = {
                    displayName: siteConfig.lists[0].listName,
                    columns: listColumns,
                    list: {
                        template: "genericList",
                    },
                };

                this.graphEndpoint = graphConfig.sitesGraphEndpoint + "/" + siteId + graphConfig.listsGraphEndpoint;

                const listCreationRes = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, listSchema);

                resolve(listCreationRes);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateAssessmentList \n",
                    JSON.stringify(ex)
                );
                reject(ex);
            }
        });
    }

    // loop through selected roles and create tag object
    private async createTagObject(teamInfo: any): Promise<any> {
        let roles: any = this.state.roleAssignments;
        roles.push({
            role: constants.incidentCommanderRoleName,
            userNamesString: this.state.createIncDetailsItem.incidentCommander.userName,
            userDetailsObj: [this.state.createIncDetailsItem.incidentCommander]
        })
        let result: any = {
            isFullyCreated: false,
            isPartiallyCreated: false,
            failedEntries: [],
            successEntries: []
        };

        return new Promise(async (resolve, reject) => {
            let allDone = false;
            let counter = 0;

            if (roles.length > 0) {
                while (!allDone) {
                    let role = roles[counter];
                    try {
                        this.graphEndpoint = graphConfig.betaGraphEndpoint + teamInfo.data.id + graphConfig.tagsGraphEndpoint;
                        const members: any = [];
                        role.userDetailsObj.forEach((users: any) => {
                            members.push({
                                "userId": users.userId
                            })
                        });
                        const tagObj = {
                            "displayName": role.role,
                            "members": members
                        }
                        let createdTag = await this.createTag(this.graphEndpoint, tagObj)

                        if (createdTag && createdTag.status) {
                            // set tag object
                            let tagCreationObj: any = {
                                tagName: role.role,
                                isCreated: true
                            };
                            result.successEntries.push(tagCreationObj);
                        }
                        else {
                            // set tag object
                            let tagCreationObj: any = {
                                tagName: role.role,
                                isCreated: false
                            };
                            result.failedEntries.push(tagCreationObj);
                        }
                        counter++;
                    } catch (ex: any) {
                        console.error(
                            constants.errorLogPrefix + "CreateIncident_CreateTag \n",
                            JSON.stringify(ex)
                        );

                        let tagCreationObj: any = {
                            tagName: role.role,
                            isCreated: false,
                            rawData: ex.message
                        };
                        result.isFullyCreated = false;
                        result.failedEntries.push(tagCreationObj);
                        counter++;

                    }
                    allDone = roles.length === counter;
                }
            }
            result.isFullyCreated = result.failedEntries.length === 0 ? true : false;
            resolve(result);
        });
    }

    // create tags for selected roles
    private async createTag(graphEndpoint: string, tagObj: any): Promise<any> {
        return new Promise(async (resolve) => {
            let maxTagCreationAttempt = 5, isTagCreated = false;

            let result = {
                status: false,
                data: {}
            };

            // loop till the tag is created
            // attempting multiple times as sometimes teams group doesn't reflect immediately after creation
            while (isTagCreated === false && maxTagCreationAttempt > 0) {
                try {
                    // logging date time stamp for debug
                    console.log(new Date());
                    // call method to create tag
                    let tagCreationInfo = await this.dataService.sendGraphPostRequest(graphEndpoint, this.props.graph, tagObj);

                    // update the result object
                    if (tagCreationInfo) {
                        isTagCreated = true;
                        result.data = tagCreationInfo;
                        result.status = true;
                    }
                } catch (creationError: any) {
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_CreateTag \n" + new Date() + "\n",
                        JSON.stringify(creationError)
                    );
                    if (creationError.statusCode === 409 && creationError.message === "Tag already exists") {
                        isTagCreated = true;
                    }
                    result.status = false;
                }
                maxTagCreationAttempt--;
            }
            console.log(constants.infoLogPrefix + "createTag_No Of Attempt", (5 - maxTagCreationAttempt), result);
            resolve(result);
        });
    }

    // method to delete team group
    private async deleteTeamGroup(group_id: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + group_id;
            try {
                const deleteResult = await this.dataService.sendGraphDeleteRequest(this.graphEndpoint, this.props.graph);
                resolve(deleteResult);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_DeleteTeamGroup \n",
                    JSON.stringify(ex)
                );
                reject(ex);
            }
        });
    }

    // method to delete created incident
    private async deleteIncident(incidentId: number): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${incidentId}`;
            try {
                const deleteResult = await this.dataService.sendGraphDeleteRequest(this.graphEndpoint, this.props.graph);
                resolve(deleteResult);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_DeleteIncident \n",
                    JSON.stringify(ex)
                );
                reject(ex);
            }
        });
    }

    // move focus to top of page to show loader or message bar
    private scrollToTop = () => {
        window.scrollTo({
            top: 0,
            behavior: 'auto'
        });
    };

    public render() {
        return (
            <>
                <div className="incident-details">

                    <>
                        {this.state.showLoader &&
                            <div className="loader-bg">
                                <div className="loaderStyle">
                                    <Loader label={this.state.loaderMessage} size="largest" />
                                </div>
                            </div>
                        }
                        <div style={{ opacity: this.state.formOpacity }}>
                            <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="incident-details-path">
                                <label>
                                    <span onClick={() => this.props.onBackClick(false)} className="go-back">
                                        <ChevronStartIcon id="path-back-icon" />
                                        <span className="back-label" title="Back">Back</span>
                                    </span> &nbsp;&nbsp;
                                    <span className="right-border">|</span>
                                    <span>&nbsp;&nbsp;{this.props.localeStrings.formTitle}</span>
                                </label>
                            </div>
                            <div className="incident-details-form-area">
                                <div className="container">
                                    <div className="incident-form-head-text">{this.props.localeStrings.formTitle}</div>
                                    <Row xs={1} sm={2} md={3}>
                                        <Col md={4} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <FormInput
                                                    label={this.props.localeStrings.fieldIncidentName}
                                                    type="text"
                                                    placeholder={this.props.localeStrings.phIncidentName}
                                                    fluid={true}
                                                    maxLength={constants.maxCharLengthForSingleLine}
                                                    required
                                                    onChange={(evt) => this.onTextInputChange(evt, "incidentName")}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.incidentName ? this.state.createIncDetailsItem.incidentName : '') : ''}
                                                    className="incident-details-input-field"
                                                    successIndicator={false}
                                                />
                                                {this.state.inputValidation.incidentNameHasError && (
                                                    <label className="message-label">{this.props.localeStrings.incidentNameRequired}</label>
                                                )}
                                                {this.state.inputRegexValidation.incidentNameHasError && (
                                                    <label className="message-label">{this.props.localeStrings.incidentNameRegex}</label>
                                                )}
                                            </div>
                                            <div className="incident-grid-item">
                                                <FormDropdown
                                                    label={{ content: this.props.localeStrings.fieldIncidentType, required: true }}
                                                    placeholder={this.props.localeStrings.phIncidentType}
                                                    items={this.state.dropdownOptions ? this.state.dropdownOptions["typeOptions"] : []}
                                                    fluid={true}
                                                    search
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.incidentType ? this.state.createIncDetailsItem.incidentType : '') : ''}
                                                    onChange={this.onIncidentTypeChange}
                                                    className="incident-type-dropdown"
                                                />
                                                {this.state.inputValidation.incidentTypeHasError && (
                                                    <label className="message-label">{this.props.localeStrings.incidentTypeRequired}</label>
                                                )}
                                            </div>
                                            <div className="incident-grid-item">

                                                <FormInput
                                                    label={this.props.localeStrings.fieldStartDate}
                                                    type="datetime-local"
                                                    placeholder={this.props.localeStrings.phStartDate}
                                                    fluid={true}
                                                    required
                                                    onChange={(evt) => this.onTextInputChange(evt, "startDateTime")}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.startDateTime ? this.state.createIncDetailsItem.startDateTime : '') : ''}
                                                    className={this.state.createIncDetailsItem && this.state.createIncDetailsItem.startDateTime ? "incident-details-date-field" : "dte-ph"}
                                                    successIndicator={false}
                                                />
                                                {this.state.inputValidation.incidentStartDateTimeHasError && (
                                                    <label className="message-label">{this.props.localeStrings.startDateRequired}</label>
                                                )}
                                            </div>
                                        </Col>
                                        <Col md={4} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <FormDropdown
                                                    label={{ content: this.props.localeStrings.fieldIncidentStatus, required: true }}
                                                    placeholder={this.props.localeStrings.phIncidentStatus}
                                                    items={this.state.dropdownOptions ? this.state.dropdownOptions["statusOptions"] : []}
                                                    fluid={true}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.incidentStatus ? this.state.createIncDetailsItem.incidentStatus : '') : ''}
                                                    onChange={this.onIncidentStatusChange}
                                                    className={this.state.createIncDetailsItem && this.state.createIncDetailsItem.incidentStatus ? "incident-details-dropdown" : "dropdown-placeholder"}
                                                />
                                                {this.state.inputValidation.incidentStatusHasError && (
                                                    <label className="message-label">{this.props.localeStrings.statusRequired}</label>
                                                )}
                                            </div>
                                            <div className="incident-grid-item">
                                                <label className="people-picker-label">{this.props.localeStrings.fieldIncidentCommander}</label>
                                                <TooltipHost
                                                    content={this.props.localeStrings.infoIncCommander}
                                                    calloutProps={calloutProps}
                                                    styles={hostStyles}
                                                >
                                                    <Icon aria-label="Info" iconName="Info" className="incCommanderInfoIcon" />
                                                </TooltipHost>
                                                <PeoplePicker
                                                    title={this.props.localeStrings.fieldIncidentCommander}
                                                    selectionMode="single"
                                                    selectionChanged={this.handleIncCommanderChange}
                                                    placeholder={this.props.localeStrings.phIncidentCommander}
                                                    className="incident-details-people-picker"
                                                />

                                                {this.state.inputValidation.incidentCommandarHasError && (
                                                    <label className="message-label">{this.props.localeStrings.incidentCommanderRequired}</label>
                                                )}
                                            </div>
                                            <div className="incident-grid-item">
                                                <FormInput
                                                    label={this.props.localeStrings.fieldLocation}
                                                    placeholder={this.props.localeStrings.phLocation}
                                                    fluid={true}
                                                    maxLength={constants.maxCharLengthForSingleLine}
                                                    required
                                                    onChange={(evt) => this.onTextInputChange(evt, "location")}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.location ? this.state.createIncDetailsItem.location : '') : ''}
                                                    className="incident-details-input-field"
                                                    successIndicator={false}
                                                />
                                                {this.state.inputValidation.incidentLocationHasError && (
                                                    <label className="message-label">{this.props.localeStrings.locationRequired}</label>
                                                )}
                                                {this.state.inputRegexValidation.incidentLocationHasError && (
                                                    <label className="message-label">{this.props.localeStrings.locationRegex}</label>
                                                )}
                                            </div>
                                        </Col>
                                        <Col md={4} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <FormTextArea
                                                    label={{ content: this.props.localeStrings.fieldDescription, required: true }}
                                                    placeholder={this.props.localeStrings.phDescription}
                                                    fluid={true}
                                                    maxLength={constants.maxCharLengthForMultiLine}
                                                    onChange={(evt) => this.onTextInputChange(evt, "incidentDesc")}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.incidentDesc ? this.state.createIncDetailsItem.incidentDesc : '') : ''}
                                                    className="incident-details-description-area"
                                                />
                                                {this.state.inputValidation.incidentDescriptionHasError && (
                                                    <label className="message-label">{this.props.localeStrings.incidentDescRequired}</label>
                                                )}
                                            </div>
                                        </Col>
                                    </Row>
                                    <div className="incident-form-head-text">{this.props.localeStrings.headerRoleAssignment}</div>
                                    <Row xs={1} sm={1} md={2}>
                                        <Col md={6} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <FormDropdown
                                                    label={this.props.localeStrings.fieldAdditionalRoles}
                                                    placeholder={this.props.localeStrings.phRoles}
                                                    items={this.state.dropdownOptions ? this.state.dropdownOptions["roleOptions"] : []}
                                                    fluid={true}
                                                    onChange={this.onRoleChange}
                                                    value={this.state.createIncDetailsItem ? (this.state.createIncDetailsItem.selectedRole ? this.state.createIncDetailsItem.selectedRole : '') : ''}
                                                    className={this.state.createIncDetailsItem && this.state.createIncDetailsItem.selectedRole ? "incident-details-dropdown" : "dropdown-placeholder"}
                                                />
                                            </div>
                                            {this.state.createIncDetailsItem.selectedRole && this.state.createIncDetailsItem.selectedRole.indexOf("New Role") > -1 ?
                                                <>
                                                    <div className="incident-grid-item">
                                                        <FormInput
                                                            label={this.props.localeStrings.fieldAddRoleName}
                                                            placeholder={this.props.localeStrings.phAddRoleName}
                                                            fluid={true}
                                                            maxLength={constants.maxCharLengthForSingleLine}
                                                            onChange={(evt) => this.onAddNewRoleChange(evt)}
                                                            value={this.state.newRoleString}
                                                            className="incident-details-input-field"
                                                            successIndicator={false}
                                                        />
                                                    </div>
                                                    <div className="incident-grid-item">
                                                        <Button
                                                            primary
                                                            onClick={this.addNewRole}
                                                            disabled={this.state.isCreateNewRoleBtnDisabled}
                                                            id={this.state.isCreateNewRoleBtnDisabled ? "manage-role-disabled-btn" : "manage-role-btn"}
                                                            fluid={!this.state.isDesktop}
                                                            title={this.props.localeStrings.btnCreateRole}
                                                        >
                                                            <img src={require("../assets/Images/AddIcon.svg").default}
                                                                alt="add"
                                                                className="manage-role-btn-icon"
                                                            />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <label className="manage-role-btn-label">{this.props.localeStrings.btnCreateRole}</label>
                                                        </Button>
                                                    </div>
                                                </>
                                                :
                                                <>
                                                    <div className="incident-grid-item">
                                                        <label className="people-picker-label">{this.props.localeStrings.fieldSearchUser}</label>
                                                        <PeoplePicker
                                                            selectionMode="multiple"
                                                            selectionChanged={this.handleAssignedUserChange}
                                                            placeholder={this.props.localeStrings.phSearchUser}
                                                            className="incident-details-people-picker"
                                                            selectedPeople={this.state.selectedUsers}
                                                        />
                                                    </div>
                                                    <div className="incident-grid-item">
                                                        <Button
                                                            primary
                                                            onClick={this.addRoleAssignment}
                                                            disabled={this.state.isAddRoleAssignmentBtnDisabled}
                                                            id={this.state.isAddRoleAssignmentBtnDisabled ? "manage-role-disabled-btn" : "manage-role-btn"}
                                                            fluid={!this.state.isDesktop}
                                                            title={this.props.localeStrings.btnAddUser}>
                                                            <img src={require("../assets/Images/AddIcon.svg").default}
                                                                alt="add"
                                                                className="manage-role-btn-icon"
                                                            />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <label className="manage-role-btn-label">{this.props.localeStrings.btnAddUser}</label>
                                                        </Button>
                                                    </div>
                                                </>
                                            }
                                        </Col>
                                        <Col md={6} sm={8} xs={12}>
                                            <div className="role-assignment-table">
                                                <Row id="role-grid-thead" xs={3} sm={3} md={3}>
                                                    <Col md={5} sm={4} xs={4} key={0}>{this.props.localeStrings.headerRole}</Col>
                                                    <Col md={5} sm={4} xs={4} key={1} className="thead-border-left">{this.props.localeStrings.headerUsers}</Col>
                                                    <Col md={2} sm={4} xs={4} key={2} className="thead-border-left col-center">{this.props.localeStrings.headerDelete}</Col>
                                                </Row>
                                                {this.state.roleAssignments.map((item, index) => (
                                                    <Row xs={3} sm={3} md={3} key={index} id="role-grid-tbody">
                                                        <Col md={5} sm={4} xs={4}>{item.role}</Col>
                                                        <Col md={5} sm={4} xs={4}>{item.userNamesString}</Col>
                                                        <Col md={2} sm={4} xs={4} className="col-center">
                                                            <img
                                                                src={require("../assets/Images/DeleteIcon.svg").default}
                                                                alt="Delete Icon"
                                                                className="role-delete-icon"
                                                                onClick={(e) => this.deleteRoleItem(index)}
                                                                title={this.props.localeStrings.headerDelete}
                                                            />
                                                        </Col>
                                                    </Row>
                                                ))}
                                            </div>
                                        </Col>
                                    </Row>
                                    <br />
                                    <Row xs={1} sm={1} md={1}>
                                        <Col md={12} sm={8} xs={12}>
                                            <div className="new-incident-btn-area">
                                                <Flex hAlign="end" gap="gap.large" wrap={true}>
                                                    <Button
                                                        onClick={() => this.props.onBackClick(false)}
                                                        id="new-incident-back-btn"
                                                        fluid={true}
                                                        title={this.props.localeStrings.btnBack}
                                                    >
                                                        <ChevronStartIcon /> &nbsp;
                                                        <label>{this.props.localeStrings.btnBack}</label>
                                                    </Button>
                                                    <Button
                                                        primary
                                                        onClick={this.createNewIncident}
                                                        fluid={true}
                                                        id="new-incident-create-btn"
                                                        title={this.props.localeStrings.btnCreateIncident}
                                                    >
                                                        <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt="edit icon" /> &nbsp;
                                                        <label>{this.props.localeStrings.btnCreateIncident}</label>
                                                    </Button>
                                                </Flex>
                                            </div>
                                        </Col>
                                    </Row>
                                </div>
                            </div>
                        </div>
                    </>

                </div>
            </>
        );
    }
}

export default IncidentDetails;
