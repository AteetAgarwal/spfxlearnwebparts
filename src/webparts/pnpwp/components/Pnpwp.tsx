import * as React from 'react';
import styles from './Pnpwp.module.scss';
import { IPnpwpProps } from './IPnpwpProps';
import { IPnpwpState } from './IPnpwpState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPOperations } from '../../Services/SPServices';
import {Button, Dropdown} from 'office-ui-fabric-react';

export default class Pnpwp extends React.Component<IPnpwpProps, IPnpwpState, {}> {
  private _spService:SPOperations;
  private selectedPnPListTitle:string;
  constructor(props:IPnpwpProps){
      super(props);
      this.state={
        listTitles:[],
        status:""
      };
      this._spService=new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<IPnpwpProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    let msg:string="";
    if(hasTeamsContext){
      msg="Welcome to SPFx teams tab! Mr. ";
    }
    else{
      msg="Welcome to SPFx webpart! Mr. ";
    }

    return (
      <section className={`${styles.pnpwp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`${styles.welcome} ${styles.row}`}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>{msg} {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Demo of PnP: <strong>{escape(description)}</strong></div>
          <div id="controls" className={styles.controls}>
            <Dropdown className={styles.dropdown} 
              options={this.state.listTitles}
              placeholder='*****Select List*****'
              onChange={this.getSelectedListTitle}>
            </Dropdown>
            <Button className={styles.mybutton} text="Create List Item" onClick={()=>this._spService.CreateListItemsByPnP(this.props.context,this.selectedPnPListTitle)
              .then((result)=>{
                this.setState({status:result})
              })}>
            </Button>
            <Button className={styles.mybutton} text="Update List Item" onClick={()=>this._spService.UpdateListItemByPnP(this.props.context,this.selectedPnPListTitle)
              .then((result)=>{
                this.setState({status:result})
              })}>
            </Button>
            <Button className={styles.mybutton} text="Delete List Item" onClick={()=>this._spService.DeleteListItemByPnP(this.props.context,this.selectedPnPListTitle)
              .then((result)=>{
                this.setState({status:result})
              })}>
            </Button>
            <div className={styles.myStatusBar}>{this.state.status}</div>
          </div>
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    this._spService.getAllListsByPnP(this.props.context).then((result)=>{
      this.setState({listTitles:result});
    })
  }

  public getSelectedListTitle=(event:any,data:any)=>{
    this.selectedPnPListTitle=data.text;
  }
}
