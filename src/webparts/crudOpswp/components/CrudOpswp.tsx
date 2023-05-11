import * as React from 'react';
import styles from './CrudOpswp.module.scss';
import { ICrudOpswpProps } from './ICrudOpswpProps';
import { ICrudOpswpState } from './ICrudOpswlState';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPOperations} from "../../Services/SPServices";
import {Button, Dropdown, IDropdownOption} from 'office-ui-fabric-react';

export default class CrudOpswp extends React.Component<ICrudOpswpProps, ICrudOpswpState,{}> {
  public _spOps: SPOperations;
  public selectedListTitle:string;
   constructor(props: ICrudOpswpProps){
    super(props);
    this._spOps = new SPOperations(this.props.context);
    this.state={
      listTitles:[],
      status:""
    }
   }
  public render(): React.ReactElement<ICrudOpswpProps> {
    const {
      description,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    return (
      <section className={`${styles.crudOpswp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`${styles.row} ${styles.row}`} >
        {environmentMessage}
        {description}
          <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
            <div className={styles.column}>
              <h3>Demo : SharePoint CRUD Operations using Rest API (spHTTPClient)</h3>
            </div>
            <div id="controls" className={styles.controls}>
              <Dropdown 
                className={styles.dropdown} 
                options={this.state.listTitles}
                placeholder='*****Select List*****'
                onChange={this.getListTitle}>
              </Dropdown>
              <Button text="Create List Item" className={styles.mybutton}
                onClick={()=>this._spOps.CreateListItems(this.props.context,this.selectedListTitle)
                  .then((result:string)=>{
                    this.setState({status:result});
              })}></Button>
              <Button text="Update List Item" className={styles.mybutton} onClick={()=>this._spOps.UpdateListItems(this.props.context,this.selectedListTitle)
                  .then((result:string)=>{
                    this.setState({status:result});})}></Button>
              <Button text="Delete List Item" className={styles.mybutton} onClick={()=>this._spOps.DeleteListItems(this.props.context,this.selectedListTitle)
                  .then((result:string)=>{
                    this.setState({status:result});})}
              ></Button>
              <div className={styles.myStatusBar}> {this.state.status}</div>
            </div>
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    this._spOps.getAllLists(this.props.context).then((result:IDropdownOption[])=>{
      this.setState({listTitles:result});
    });
  }

  public getListTitle=(event:any, data:any)=>{
    this.selectedListTitle=data.text;
  }
}
