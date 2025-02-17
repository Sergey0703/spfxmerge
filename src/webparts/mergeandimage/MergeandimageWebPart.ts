import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IMergeandimageProps } from './components/IMergeandimageProps';
import Mergeandimage from './components/Mergeandimage';

export default class MergeandimageWebPart extends BaseClientSideWebPart<IMergeandimageProps> {

  private _environmentMessage: string = "Hello from the web part!";
  
  public render(): void {
    this.context.msGraphClientFactory.getClient("3")
      .then((graphClient: MSGraphClientV3) => {
        const element: React.ReactElement<IMergeandimageProps> = React.createElement(
          Mergeandimage,
          {
            description: this.properties.description,
            graphClient: graphClient,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: true,  // Set this to true or false based on your requirements
            userDisplayName: this.context.pageContext.user.displayName
          }
        );
        ReactDom.render(element, this.domElement);
      })
      .catch((error) => {
        console.error("Error initializing MSGraphClient: ", error);
      });
  }

  // Optional lifecycle methods can be added if necessary
}
