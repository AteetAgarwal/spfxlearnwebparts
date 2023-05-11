import * as React from 'react';
import styles from './Handlelargelist.module.scss';
import { IHandlelargelistProps } from './IHandlelargelistProps';
import { IHandlelargelistState } from './IHandlelargelistState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList } from 'office-ui-fabric-react';
import { SPOperations } from '../../Services/SPServices';
import { IListItem } from './IListItem';

export default class Handlelargelist extends React.Component<IHandlelargelistProps,IHandlelargelistState, {}> {
  private _spServices: SPOperations;
  constructor(props:IHandlelargelistProps){
    super(props);
    this.state={
      listResults:[]
    };
    this._spServices=new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<IHandlelargelistProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    return (
      <section className={`${styles.handlelargelist} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          Total Results: {this.state.listResults.length}
          <DetailsList items={this.state.listResults}
          onShouldVirtualize={ () => false }>
          </DetailsList>
        </div>
      </section>
    );
  }

  public componentDidMount(): void {
    /*
    //PnP function calls
    this._spServices.getMoreThan5KListItemsWithoutWhereClause(this.props.listName).then((results:IListItem[])=>{
      this.setState({listResults:results});
      console.log("Total results: "+ results.length);
    });
    this._spServices.getMoreThan5KListItemsWithWhereClause(this.props.listName).then((results:IListItem[])=>{
      this.setState({listResults:results});
      console.log("Total results: "+ results.length);
    });*/
    
    /*
    //SPHttpClient calls using Rest API
    this._spServices.getMoreThan5KListItemsRestAPIWithWhereClause(this.props.context,this.props.listName).then((results:IListItem[])=>{
      this.setState({listResults:results});
      console.log("Total results: "+ results.length);
    });*/
    this._spServices.getMoreThan5KListItemsRestAPIWithoutWhereClause(this.props.context,this.props.listName).then((results:IListItem[])=>{
      this.setState({listResults:results});
      console.log("Total results: "+ results.length);
    });
  }
}
