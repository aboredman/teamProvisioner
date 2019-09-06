import * as React from "react";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext
} from "msteams-ui-components-react";
import { Dropdown, DropdownMenuItemType, IDropdownOption, TextField } from 'office-ui-fabric-react';
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the teamProvisionerTabApproveTab React component
 */
export interface ITeamProvisionerTabApproveState extends ITeamsBaseComponentState {
    entityId?: string;
    reqTeamName?: string;
    reqTeamType?: { key: string | number | undefined };
    isUserAdmin?: boolean;
}

/**
 * Properties for the teamProvisionerTabApproveTab React component
 */
export interface ITeamProvisionerTabApproveProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the TeamProvisioner content page
 */
export class TeamProvisionerTabApprove extends TeamsBaseComponent<ITeamProvisionerTabApproveProps, ITeamProvisionerTabApproveState> {

    public componentWillMount() {
        
        
        
        
        
        
        
        
        
        
        
        
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId,
                    isUserAdmin: false
                });
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Request a new team</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                {this.state.entityId}

                            </div>
                            <div style={styles.section}>
                                this is the approve tab
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>


                                (C) Copyright XYZ
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
}
