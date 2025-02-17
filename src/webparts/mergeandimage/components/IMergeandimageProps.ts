import { MSGraphClientV3 } from '@microsoft/sp-http'; // Import this to resolve MSGraphClientV3

export interface IMergeandimageProps {
  description: string;
  graphClient: MSGraphClientV3;
  environmentMessage?: string; // Optional, can be set to undefined if not needed
  hasTeamsContext?: boolean; // Optional
  userDisplayName?: string; // Optional
}
