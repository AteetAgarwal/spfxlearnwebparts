import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IList } from "../../sourceconnectecwebpart/SourceconnectecwebpartWebPart";
import {DynamicProperty} from "@microsoft/sp-component-base";

export interface IConsumerconnectedwebpartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  ListTitle: DynamicProperty<IList>;
}
