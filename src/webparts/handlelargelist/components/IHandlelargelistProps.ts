import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHandlelargelistProps {
  description: string;
  context: WebPartContext;
  listName:string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
