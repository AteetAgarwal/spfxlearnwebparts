import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IElevatepermissionProps {
  description: string;
  context: WebPartContext;
  flowUrl:string;
  listTitle:string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
