import * as React from 'react';
import styles from './FileTeamsSaverAddin.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { Stack } from 'office-ui-fabric-react';

import { ServiceScope } from "@microsoft/sp-core-library";
import { TeamsService } from '../../../services/TeamsService';

import Tree from '@naisutech/react-tree';
import INode from '../../../model/INode';

export interface IFileTeamsSaverAddinProps {
  serviceScope: ServiceScope,
  mailId: string;
}

export interface IFileTeamsSaverAddinState {
  dataLoaded: boolean;
  nodes: INode[];
  selectedNode: INode;
  saveButtonEnabled: boolean;
}

export default class FileTeamsSaverAddin extends React.Component<IFileTeamsSaverAddinProps, IFileTeamsSaverAddinState> {

  constructor(props: IFileTeamsSaverAddinProps) {
    super(props);
    this.teamsService = props.serviceScope.consume(TeamsService.serviceKey);
    this.state = {
      dataLoaded: false,
      nodes: [],
      selectedNode: null,
      saveButtonEnabled: false
    };
  }

  private teamsService: TeamsService;

  public componentDidMount() {
    let nodes: INode[] = [{ parentId: null, id: "0", label: "Teams", selectable: false }];
    this.teamsService.getMyTeams().then(teams => {
      Promise.all(
        teams.map(team => {
          nodes.push(team);
          return this.teamsService.getChannels(team.id).then(channels => {
            channels.map(channel => nodes.push(channel));
          });
        }))
        .then(() => {
          this.setState({
            dataLoaded: true,
            nodes: this.state.nodes.concat(nodes)
          });
        });
    });
  }

  private saveClick = (): void => {
    this.teamsService.getAttachments(this.props.mailId).then(response => {
      response.map(attachment => {
        this.teamsService.getRawAttachment(this.props.mailId, attachment.id).then(rawAttachment => {
          this.teamsService.uploadFile(rawAttachment, this.state.selectedNode.parentId, this.state.selectedNode.id, attachment.name, attachment.contentType);
        });
      });
    });
  };

  private selectNode = (node: INode): void => {
    this.setState({ ...this.state, selectedNode: node, saveButtonEnabled: node.selectable });
  }

  public render(): React.ReactElement<IFileTeamsSaverAddinProps> {
    return (
      <div>
        <Stack horizontalAlign="center">
          {this.state.dataLoaded
            ? <React.Fragment>
              <Tree nodes={this.state.nodes} onSelect={this.selectNode} theme={'light'} />
              <PrimaryButton text="Save" onClick={this.saveClick} disabled={!this.state.saveButtonEnabled} />
            </React.Fragment>
            : <Spinner size={SpinnerSize.large} />
          }
        </Stack>
      </div >
    );
  }
}
