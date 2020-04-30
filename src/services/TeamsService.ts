import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

import INode from "../model/INode";


export class TeamsService {
    public static readonly serviceKey: ServiceKey<TeamsService> =
        ServiceKey.create('TeamsService', TeamsService);

    private _graphClient: MSGraphClient;
    private _graphClientFactory: MSGraphClientFactory

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }

    public getMyTeams(): Promise<INode[]> {
        return this.getClient().then(client => {
            return client.api("/me/joinedTeams").get().then(response => {
                let teams: [MicrosoftGraph.Team & MicrosoftGraph.Group] = response.value;
                return teams.map(team => {
                    return {
                        id: team.id,
                        label: team.displayName,
                        parentId: "0",
                        selectable: false
                    } as INode
                });
            });
        });
    }

    public getChannels(id: string): Promise<INode[]> {
        return this._graphClient.api(`https://graph.microsoft.com/v1.0/groups/${id}/drive/root/children`).get().then(response => {
            let channels: [MicrosoftGraph.BaseItem] = response.value;
            return channels.map(channel => {
                return {
                    id: channel.id,
                    label: channel.name,
                    parentId: id,
                    selectable: true
                } as INode
            });
        });
    }

    public getAttachments(messageId: string): Promise<any> {
        return this._graphClient.api(`https://graph.microsoft.com/v1.0/me/messages/${messageId.replace("/", "-")}/attachments/`).get()
            .then(response => {
                let attachments: MicrosoftGraph.FileAttachment = response.value;
                return attachments;
            });
    }


    public getRawAttachment(messageId: string, attachmentId: string): Promise<any> {
        return this._graphClient.api(`https://graph.microsoft.com/v1.0/me/messages/${messageId.replace("/", "-")}/attachments/${attachmentId}/$value`)
            .responseType('ARRAYBUFFER')
            .get()
            .then(rawAttachment => {
                return rawAttachment;
            });
    }

    public uploadFile(attachment: any, groupId: string, channelId: string, fileName: string, contentType: string): Promise<void> {
        return this._graphClient.api(`https://graph.microsoft.com/v1.0/groups/${groupId}/drive/items/${channelId}:/${fileName}:/content`)
            .header("Content-Type", contentType)
            .put(attachment)
            .then(response => {
                console.log(response);
            });
    }


    private getClient(): Promise<MSGraphClient> {
        return this._graphClient
            ? Promise.resolve(this._graphClient)
            : this._graphClientFactory.getClient().then(graphClient => {
                this._graphClient = graphClient;
                return graphClient;
            })
    }
}