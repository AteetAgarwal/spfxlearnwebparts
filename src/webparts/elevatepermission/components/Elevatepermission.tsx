import * as React from 'react';
import styles from './Elevatepermission.module.scss';
import { IElevatepermissionProps } from './IElevatepermissionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IElevatepermissionState } from './IElevatepermissionState';
import { Button } from 'office-ui-fabric-react';
import { SPOperations } from '../../Services/SPServices';

export default class Elevatepermission extends React.Component<IElevatepermissionProps, IElevatepermissionState, {}> {
  private _spService:SPOperations;
  constructor(props:IElevatepermissionProps){
    super(props);
      this.state={
        status:""
      };
      this._spService=new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<IElevatepermissionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.elevatepermission} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div className={styles.controls}>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <Button className={styles.mybutton} text="Create Item" onClick={()=> this.setState({status:"In Progress"},()=>{
                        this._spService.CallPowerAutomate(this.props.context,this.props.listTitle, this.props.flowUrl)
                        .then((result)=>{
                          this.setState({status:"Item created successfully"})
                        })
          })}>
          </Button>
          <div id="dv_Status" className={styles.myStatusBar}>{this.state.status}</div>
        </div>
      </section>
    );
  }
}
