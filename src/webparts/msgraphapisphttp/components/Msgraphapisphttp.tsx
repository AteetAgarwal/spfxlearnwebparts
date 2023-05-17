import * as React from 'react';
//import styles from './Msgraphapisphttp.module.scss';
import { IMsgraphapisphttpProps } from './IMsgraphapisphttpProps';
import { DetailsList, PrimaryButton } from 'office-ui-fabric-react';
//import { escape } from '@microsoft/sp-lodash-subset';
import {SPOperations} from '../../Services/SPServices';

export interface IUser{
  displayName: string;
  mail:string;
}

export interface IUserState{
  users:IUser[];
  items:IList[];
}

export interface IList{
  title:string;
  mail:string;
}

export default class Msgraphapisphttp extends React.Component<IMsgraphapisphttpProps,IUserState, {}> {
  private _spServices: SPOperations;

  constructor(props:IMsgraphapisphttpProps){
    super(props);
    this.state={
      users:[],
      items:[]
    };
    this._spServices=new SPOperations(this.props.context);
  }
  
  public render(): React.ReactElement<IMsgraphapisphttpProps> {
    return (
      <div>
        <PrimaryButton text='Search Users'
          onClick={this.getUsers}></PrimaryButton>
        <DetailsList items={this.state.users} onShouldVirtualize={ () => false }  ></DetailsList>
        <br/><br/>
        <PrimaryButton text='Get Items From List'
          onClick={this.getListData}></PrimaryButton>
        <DetailsList items={this.state.items} onShouldVirtualize={ () => false }  ></DetailsList>
      </div>
    );
  }

  public getUsers=():void=>{
    this._spServices.getUsers(this.props.context).then((response)=>{
      this.setState({users:response});
    });
  }

  public getListData=():void=>{
    this._spServices.getAllItemsFromList(this.props.context, "SPMockData").then((response)=>{
      this.setState({items:response});
    });
  }
}
