import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IList } from "../SourceconnectecwebpartWebPart";

export interface ISourceconnectecwebpartProps {
  description: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  PassListTitle: (title:IList)=>void;
  environmentMessage: string;
}
