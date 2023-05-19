import * as React from 'react';
//import styles from './Msgraphapipnpwp.module.scss';
import { IMsgraphapipnpwpProps } from './IMsgraphapipnpwpProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPOperations } from '../../Services/SPServices';

export default class Msgraphapipnpwp extends React.Component<IMsgraphapipnpwpProps, {}> {
  private spServices: SPOperations;
  constructor(props:IMsgraphapipnpwpProps){
    super(props);
    this.spServices = new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<IMsgraphapipnpwpProps> {
    return (
      <div id="parent">Calendar Events:</div>
    );
  }

  public componentDidMount(): void {
    this.spServices.getCalendarEvents(this.props.context).then((response)=>{
      console.log(response);
    });
  }
}
