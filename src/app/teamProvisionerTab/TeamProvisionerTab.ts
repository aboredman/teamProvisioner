import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/teamProvisionerTab/index.html")
@PreventIframe("/teamProvisionerTab/config.html")
@PreventIframe("/teamProvisionerTab/remove.html")
@PreventIframe("/teamProvisionerTab/approve.html")

export class TeamProvisionerTab {
}
