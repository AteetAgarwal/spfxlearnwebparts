import {WebPartContext} from "@microsoft/sp-webpart-base";
export interface ICrudOpswpProps {
  description: string;
  context: WebPartContext;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
