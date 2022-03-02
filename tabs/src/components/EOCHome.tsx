import React from 'react';
import { TeamsUserCredential, createMicrosoftGraphClient } from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";
import { Providers, ProviderState, SimpleProvider } from '@microsoft/mgt-element';
import { Button, Loader } from "@fluentui/react-northstar";
import { MessageBar, MessageBarType, initializeIcons } from '@fluentui/react';
import Dashboard from './Dashboard';
import EocHeader from './EocHeader';
import siteConfig from '../config/siteConfig.json';
import * as graphConfig from '../common/graphConfig';
import CommonService from "../common/CommonService";
import * as constants from '../common/Constants';
import IncidentDetails from './IncidentDetails';
import "../scss/EOCHome.module.scss";
import * as microsoftTeams from "@microsoft/teams-js";
import LocalizedStrings from 'react-localization';
import { localizedStrings } from "../locale/LocaleStrings";

initializeIcons();

interface IEOCHomeState {
    showLoginPage: boolean;
    graph: Client;
    tenantName: string;
    siteId: string;
    showIncForm: boolean;
    showMessageBar: boolean;
    message: string;
    messageBarType: string;
    locale: string;
    currentUserName: string;
    currentUserId: string;
    loaderMessage: string;
}

interface IEOCHomeProps {
}

let localeStrings = new LocalizedStrings(localizedStrings);

export class EOCHome extends React.Component<IEOCHomeProps, IEOCHomeState>  {
    private credential = new TeamsUserCredential();
    private scope = graphConfig.scope;
    private dataService = new CommonService();

    constructor(props: any) {
        super(props);

        const { scope } = {
            scope: graphConfig.scope
        };

        // create graph client without asking for login based on previous sessions
        const credential = new TeamsUserCredential();
        const graph = createMicrosoftGraphClient(credential, scope);

        this.state = {
            showLoginPage: true,
            graph: graph,
            tenantName: '',
            siteId: '',
            showIncForm: false,
            showMessageBar: false,
            message: "",
            messageBarType: "",
            locale: "",
            currentUserName: "",
            currentUserId: "",
            loaderMessage: localeStrings.genericLoaderMessage,
        }
    }

    async componentDidMount() {
        await this.initGraphToolkit(new TeamsUserCredential(), graphConfig.scope);
        await this.checkIsConsentNeeded();

        try {
            // get current user's language from Teams App settings
            microsoftTeams.getContext(ctx => {
                if (ctx && ctx.locale && ctx.locale !== "") {
                    this.setState({
                        locale: ctx.locale
                    })
                }
                else {
                    this.setState({
                        locale: constants.defaultLocale
                    })
                }
            })
        } catch (error) {
            this.setState({
                locale: constants.defaultLocale
            })
        }

        // call method to get the tenant details
        if (!this.state.showLoginPage) {
            await this.getTenantAndSiteDetails();
            await this.getCurrentUserDetails();
        }
    }

    // Initialize the toolkit and get access token
    async initGraphToolkit(credential: any, scopeVar: any) {

        async function getAccessToken(scopeVar: any) {
            let tokenObj = await credential.getToken(scopeVar);
            return tokenObj.token;
        }

        async function login() {
            try {
                await credential.login(scopeVar);
            } catch (err) {
                alert("Login failed: " + err);
                return;
            }
            Providers.globalProvider.setState(ProviderState.SignedIn);
        }

        async function logout() { }

        Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);
        Providers.globalProvider.setState(ProviderState.SignedIn);
    }

    // check if token is valid else show login to get token
    async checkIsConsentNeeded() {
        try {
            await this.credential.getToken(this.scope);
        } catch (error) {
            this.setState({
                showLoginPage: true
            });
            return true;
        }
        this.setState({
            showLoginPage: false
        });
        return false;
    }

    // this function gets called on Authorized button click
    public loginClick = async () => {
        const { scope } = {
            scope: graphConfig.scope
        };
        const credential = new TeamsUserCredential();
        await credential.login(scope);
        const graph = createMicrosoftGraphClient(credential, scope); // create graph object

        const profile = await this.dataService.getGraphData(graphConfig.meGraphEndpoint, this.state.graph); // get user profile to validate the API

        // validate if the above API call is returning result
        if (!!profile) {
            this.setState({ showLoginPage: false, graph: graph })

            // call method to get the tenant details
            if (!this.state.showLoginPage) {
                await this.getTenantAndSiteDetails();
                await this.getCurrentUserDetails();
            }
        }
        else {
            this.setState({ showLoginPage: true })
        }
    }

    // this method connects with service layer to get the tenant name and SharePoint site Id
    public async getTenantAndSiteDetails() {
        try {
            // get the tenant name
            const tenantName = await this.dataService.getTenantDetails(graphConfig.organizationGraphEndpoint, this.state.graph);

            // Form the graph end point to get the SharePoint site Id
            const urlForSiteId = graphConfig.spSiteGraphEndpoint + tenantName + ".sharepoint.com:/sites/" + siteConfig.siteName + "?$select=id";

            // get SharePoint site Id
            const siteDetails = await this.dataService.getGraphData(urlForSiteId, this.state.graph);

            this.setState({
                tenantName: tenantName,
                siteId: siteDetails.id
            })
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_EOCHome_GetTenantAndSiteDetails \n",
                JSON.stringify(error)
            );
        }
    }

    // this method connects with service layer to get the current user details
    public async getCurrentUserDetails() {
        try {
            // get the tenant name
            const currentUser = await this.dataService.getGraphData(graphConfig.meGraphEndpoint, this.state.graph);
            this.setState({
                currentUserName: currentUser.givenName,
                currentUserId: currentUser.id
            })
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_EOCHome_GetTenantAndSiteDetails \n",
                JSON.stringify(error)
            );
        }
    }

    // changes state to hide message bar
    private hideMessageBar = () => {
        this.setState({
            showMessageBar: false
        })
    }

    // changes state to show message bar
    private showMessageBar = (message: string, type: string) => {
        this.setState({
            showMessageBar: true,
            message: message,
            messageBarType: type
        })
    }

    // changes state to show message bar and dashboard
    private handleBackClick = (showMessageBar: boolean) => {
        if (showMessageBar) {
            this.setState({ showIncForm: false })
        }
        else {
            this.setState({ showIncForm: false, showMessageBar: false })
        }
    }

    public render() {
        // let localeStrings = new LocalizedStrings(localizedStrings);
        if (this.state.locale && this.state.locale !== "") {
            localeStrings.setLanguage(this.state.locale);
        }

        return (
            <>
            {this.state.locale === "" ?
                <>
                    <Loader label={this.state.loaderMessage} size="largest" />
                </>
                :                
                    <>
                        <EocHeader clickcallback={() => { }}
                            localeStrings={localeStrings}
                            currentUserName={this.state.currentUserName} />
                        {this.state.showLoginPage &&
                            <div className='loginButton'>
                                <Button primary content={localeStrings.btnLogin} disabled={!this.state.showLoginPage} onClick={this.loginClick} />
                            </div>
                        }
                        {!this.state.showLoginPage && this.state.siteId !== "" &&
                            <div>
                                {this.state.showMessageBar &&
                                    <>
                                        {this.state.messageBarType === "success" &&
                                            <MessageBar
                                                messageBarType={MessageBarType.success}
                                                isMultiline={false}
                                                dismissButtonAriaLabel="Close"
                                                onDismiss={this.hideMessageBar}
                                                className="message-bar"
                                            >
                                                {this.state.message}
                                            </MessageBar>
                                        }
                                        {this.state.messageBarType === "error" &&
                                            <MessageBar
                                                messageBarType={MessageBarType.error}
                                                isMultiline={false}
                                                dismissButtonAriaLabel="Close"
                                                onDismiss={this.hideMessageBar}
                                                className="message-bar"
                                            >
                                                {this.state.message}
                                            </MessageBar>
                                        }
                                    </>
                                }
                                {!this.state.showIncForm ?
                                    <Dashboard
                                        graph={this.state.graph}
                                        tenantName={this.state.tenantName}
                                        siteId={this.state.siteId}
                                        onCreateTeamClick={() => {
                                            this.setState({ showIncForm: true });
                                            this.hideMessageBar();
                                        }}
                                        localeStrings={localeStrings}
                                        showMessageBar={this.showMessageBar}
                                        hideMessageBar={this.hideMessageBar}
                                    />
                                    :
                                    <IncidentDetails
                                        graph={this.state.graph}
                                        tenantName={this.state.tenantName}
                                        siteId={this.state.siteId}
                                        onBackClick={this.handleBackClick}
                                        showMessageBar={this.showMessageBar}
                                        hideMessageBar={this.hideMessageBar}
                                        localeStrings={localeStrings}
                                        currentUserId={this.state.currentUserId}
                                    />
                                }
                            </div>
                        }
                    </>
                }
            </>
        )
    }
}
