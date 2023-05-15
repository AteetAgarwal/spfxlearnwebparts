import * as React from 'react';
import styles from './Sourceconnectecwebpart.module.scss';
import { ISourceconnectecwebpartProps } from './ISourceconnectecwebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPOperations} from "../../Services/SPServices";
import {Dropdown, IDropdownOption} from 'office-ui-fabric-react';
import { ISourceconnectecwebpartState } from './ISourceconnectecwebpartState';
import { IList } from '../SourceconnectecwebpartWebPart';

export default class Sourceconnectecwebpart extends React.Component<ISourceconnectecwebpartProps,ISourceconnectecwebpartState, {}> {
  private _spOps: SPOperations;
  public selectedListTitle:IList;
  constructor(props: ISourceconnectecwebpartProps){
    super(props);
    this._spOps = new SPOperations(this.props.context);
    this.state={
      listTitles:[],
      status:""
    }
   }
  public render(): React.ReactElement<ISourceconnectecwebpartProps> {
    const {
      description,
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.sourceconnectecwebpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div id="controls" className={styles.controls}>
              <Dropdown 
                className={styles.dropdown} 
                options={this.state.listTitles}
                placeholder='*****Select List*****'
                onChange={this.onListTitleChange}>
              </Dropdown>
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    this._spOps.getAllLists(this.props.context).then((result:IDropdownOption[])=>{
      this.setState({listTitles:result});
    });
  }

  public onListTitleChange=(ev:any,option:any):void=>{
    let listTitle:IList;
    listTitle=option.text;
    this.props.PassListTitle(listTitle);
  }
}
