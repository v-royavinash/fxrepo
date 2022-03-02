import React, { Component } from 'react';
import { Dialog, Flex, CloseIcon, FormInput, FormDropdown, SyncIcon, Button } from '@fluentui/react-northstar';
import "../scss/UpdateIncident.module.scss";
import { Col, Row } from 'react-bootstrap';
import { Client } from "@microsoft/microsoft-graph-client";
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import * as constants from '../common/Constants';
import CommonService from "../common/CommonService";
import { IInputRegexValidationStates } from '../common/CommonService';

export interface IListItem {
    itemId: string;
    incidentId: string;
    incidentName: string;
    incidentCommander: string;
    status: string;
    location: string;
    startDate: string;
    startDateUTC: string;
    createdDate: string;
}

export class IUpdateIncidentEntity {
    incidentName!: string;
    incidentStatus!: string;
    location!: string;
}

export interface IInputValidationStates {
    incidentNameHasError: boolean;
    incidentStatusHasError: boolean;
    incidentLocationHasError: boolean;
}

export interface IUpdateIncidentProps {
    openPopup: boolean;
    closePopup: (isRefreshNeeded: boolean) => void;
    incidentData: IListItem;
    graph: Client;
    tenantName: string;
    siteId: string;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    localeStrings: any;
};

export interface IUpdateIncidentState {
    isDesktop: boolean;
    statusOptions: [];
    selectedStatus: string;
    isDuplicateName: boolean;
    isDisabled: boolean;
    updateIncidentItem: IUpdateIncidentEntity;
    inputValidation: IInputValidationStates;
    inputRegexValidation: IInputRegexValidationStates;
};

// sets the initial values for required fields validation object
const getInputValildationInitialState = (): IInputValidationStates => {
    return {
        incidentNameHasError: false,
        incidentStatusHasError: false,
        incidentLocationHasError: false,
    };
};

export default class UpdateIncident extends Component<IUpdateIncidentProps, IUpdateIncidentState> {
    constructor(props: IUpdateIncidentProps) {
        super(props);
        this.state = {
            isDesktop: window.innerWidth >= constants.mobileWidth ? true : false,
            statusOptions: [],
            selectedStatus: "",
            isDuplicateName: false,
            isDisabled: false,
            updateIncidentItem: {
                incidentName: this.props.incidentData.incidentName,
                incidentStatus: this.props.incidentData.status,
                location: this.props.incidentData.location,
            },
            inputValidation: getInputValildationInitialState(),
            inputRegexValidation: this.dataService.getInputRegexValildationInitialState(),
        };
    }

    // initialize data service
    private dataService = new CommonService();
    private graphEndpoint = "";

    public async componentDidMount() {
        await this.getStatusOptions();
    }

    // get dropdown options for status
    public async getStatusOptions() {
        try {
            const incStatusGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incStatusList}/items?$expand=fields&$Top=5000`;
            const statusOptions = await this.dataService.getDropdownOptions(incStatusGraphEndpoint, this.props.graph);
            this.setState({
                statusOptions: statusOptions
            })
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_UpdateIncident_GetStatusOptions \n",
                JSON.stringify(error)
            );
        }
    }

    // on change handler for text input changes
    private onTextInputChange = (event: any, key: string) => {
        let incInfo = { ...this.state.updateIncidentItem };
        let inputValidationObj = this.state.inputValidation;
        let regexValidationObj = this.state.inputRegexValidation;
        if (incInfo) {
            switch (key) {
                case "incidentName":
                    incInfo[key] = event.target.value;

                    // check for required field validation
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentNameHasError = false;
                    }
                    else {
                        inputValidationObj.incidentNameHasError = true;
                    }
                    this.setState({
                        updateIncidentItem: incInfo,
                        inputValidation: inputValidationObj
                    })
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
                        updateIncidentItem: incInfo,
                        inputValidation: inputValidationObj
                    })
                    this.setState({ updateIncidentItem: incInfo, inputValidation: inputValidationObj })
                    break;

                default:
                    break;
            }
        }
    }

    // on incident status dropdown value change
    private onIncidentStatusChange = (event: any, selectedValue: any) => {
        let incInfo = this.state.updateIncidentItem;
        if (incInfo) {
            let incInfo = { ...this.state.updateIncidentItem };
            if (incInfo) {
                incInfo["incidentStatus"] = selectedValue.value;
                let inputValidationObj = this.state.inputValidation;
                inputValidationObj.incidentStatusHasError = false;
                this.setState({ updateIncidentItem: incInfo, inputValidation: inputValidationObj })
            }
        }
    }

    // perform required fields validation
    private requiredFieldValidation = (incidentInfo: IUpdateIncidentEntity): boolean => {
        let inputValidationObj = getInputValildationInitialState();
        let reqFieldValidationSuccess = true;
        if (incidentInfo.incidentName === "" || incidentInfo.incidentName === undefined) {
            inputValidationObj.incidentNameHasError = true;
        }
        if (incidentInfo.location === "" || incidentInfo.location === undefined) {
            inputValidationObj.incidentLocationHasError = true;
        }

        if (inputValidationObj.incidentNameHasError || inputValidationObj.incidentLocationHasError) {
            this.setState({
                inputValidation: inputValidationObj,
                isDisabled: false
            });
            reqFieldValidationSuccess = false;
        }
        return reqFieldValidationSuccess;
    }

    // on update button click
    private updateIncident = async () => {
        try {
            this.props.hideMessageBar();
            this.setState({
                inputRegexValidation: this.dataService.getInputRegexValildationInitialState(),
                inputValidation: getInputValildationInitialState(),
                isDisabled: true
            })

            // validate for required fields
            if (this.requiredFieldValidation(this.state.updateIncidentItem)) {
                // validate input strings for incident name and location
                const regexValidation = this.dataService.regexValidation(this.state.updateIncidentItem);
                if (regexValidation.incidentLocationHasError || regexValidation.incidentNameHasError) {
                    this.setState({
                        inputRegexValidation: regexValidation,
                        isDisabled: false
                    });
                }
                else {
                    let incNameStr = this.state.updateIncidentItem.incidentName.replace(/'/g, "''");
                    // create graph endpoint for querying Incident Transaction list
                    this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentsList}/items?$expand=fields&$filter=fields/IncidentName eq '${incNameStr}' and fields/IncidentId ne '${this.props.incidentData.itemId}'`;
                    // show error if incident with same name already exists
                    if (await this.dataService.getIncident(this.graphEndpoint, this.props.graph)) {
                        this.setState({
                            isDuplicateName: true,
                            isDisabled: false
                        })
                    }
                    else {
                        this.setState({
                            isDuplicateName: false
                        })
                        const updatedIncidentObject = {
                            IncidentStatus: this.state.updateIncidentItem.incidentStatus,
                            IncidentName: this.state.updateIncidentItem.incidentName,
                            Location: this.state.updateIncidentItem.location,
                        }
                        this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${this.props.incidentData.itemId}/fields`;

                        // let service = new CommonService();
                        const updatedItem = await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, updatedIncidentObject);
                        if (updatedItem) {
                            console.log(constants.infoLogPrefix + "Incident Updated");
                            this.props.closePopup(true);
                            this.props.showMessageBar(this.props.localeStrings.updateStatusSuccessMessage, constants.messageBarType.success);
                        }
                    }
                }
            }
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_UpdateIncident_UpdateStatus \n",
                JSON.stringify(error)
            );
            this.props.closePopup(true);
            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForUpdateIncident, constants.messageBarType.error);
        }
    }

    render() {
        return (
            <div>
                <Dialog
                    open={this.props.openPopup}
                    closeOnOutsideClick={false}
                    id="incident-popup"
                    content={<>
                        <Flex space="between" id="incident-popup-header">
                            <div className="popup-header-text">
                                {this.props.localeStrings.manageIncFormTitle}
                            </div>
                            <CloseIcon onClick={() => { this.props.closePopup(false); }} id="popup-header-close" />
                        </Flex>
                        <div className="incident-popup-body">
                            <Row xs={1} sm={2} md={3}>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldIncidentId}
                                            type="text"
                                            placeholder={this.props.localeStrings.phIncidentId}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.incidentId : ""}
                                            disabled
                                            className="popup-text-field-disabled"
                                        />
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormDropdown
                                            label={{ content: this.props.localeStrings.fieldIncidentStatus, required: true, className: "status-dd-label" }}
                                            placeholder={this.props.localeStrings.phIncidentStatus}
                                            items={this.state.statusOptions}
                                            fluid={true}
                                            value={this.state.updateIncidentItem ? (this.state.updateIncidentItem.incidentStatus ? this.state.updateIncidentItem.incidentStatus : '') : ''}
                                            onChange={this.onIncidentStatusChange}
                                            className="popup-dropdown"
                                        />
                                    </div>
                                </Col>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={{ content: this.props.localeStrings.fieldIncidentName }}
                                            placeholder={this.props.localeStrings.phIncidentName}
                                            fluid={true}
                                            maxLength={constants.maxCharLengthForSingleLine}
                                            value={this.state.updateIncidentItem ? this.state.updateIncidentItem.incidentName : ""}
                                            onChange={(evt) => this.onTextInputChange(evt, "incidentName")}
                                            className="popup-text-field"
                                        />
                                        {this.state.inputValidation.incidentNameHasError && (
                                            <label className="message-label">{this.props.localeStrings.incidentNameRequired}</label>
                                        )}
                                        {this.state.inputRegexValidation.incidentNameHasError && (
                                            <label className="message-label">{this.props.localeStrings.incidentNameRegex}</label>
                                        )}
                                        {this.state.isDuplicateName && (
                                            <label className="message-label">{this.props.localeStrings.duplicateIncNameOnUpdate}</label>
                                        )}
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldLocation}
                                            type="text"
                                            placeholder={this.props.localeStrings.phLocation}
                                            fluid={true}
                                            maxLength={constants.maxCharLengthForSingleLine}
                                            value={this.state.updateIncidentItem ? this.state.updateIncidentItem.location : ""}
                                            onChange={(evt) => this.onTextInputChange(evt, "location")}
                                            className="popup-text-field"
                                        />
                                        {this.state.inputValidation.incidentLocationHasError && (
                                            <label className="message-label">{this.props.localeStrings.locationRequired}</label>
                                        )}
                                        {this.state.inputRegexValidation.incidentLocationHasError && (
                                            <label className="message-label">{this.props.localeStrings.locationRegex}</label>
                                        )}
                                    </div>
                                </Col>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldIncidentCommander}
                                            type="text"
                                            placeholder={this.props.localeStrings.phIncidentCommander}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.incidentCommander : ""}
                                            disabled
                                            className="popup-text-field-disabled"
                                        />
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldStartDate}
                                            type="text"
                                            placeholder={this.props.localeStrings.phStartDate}
                                            fluid={true}
                                            defaultValue={this.props.incidentData.startDate}
                                            disabled
                                            className="popup-text-field-disabled"
                                        />
                                    </div>
                                </Col>
                            </Row>
                        </div>
                        <Flex hAlign={this.state.isDesktop ? "end" : "center"} gap="gap.small" id="popup-btn-area">
                            <Button
                                icon={<CloseIcon />}
                                content={this.props.localeStrings.btnClose}
                                iconPosition="before"
                                id="popup-close-btn"
                                title={this.props.localeStrings.btnClose}
                                onClick={() => { this.props.closePopup(false); }}
                            />
                            <Button
                                icon={<SyncIcon />}
                                content={this.props.localeStrings.btnUpdateInc}
                                iconPosition="before"
                                primary
                                disabled={this.state.isDisabled}
                                onClick={this.updateIncident}
                                id="popup-update-btn"
                                title={this.props.localeStrings.btnUpdateInc}
                                fluid={this.state.isDesktop ? false : true}
                            />
                        </Flex>
                    </>}
                />
            </div>
        )
    }
}
