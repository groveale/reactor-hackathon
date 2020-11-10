import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IGraphPersonaProps {
  graphClient: MSGraphClient;
  spfxContext: WebPartContext;
}