import * as React from 'react';
import styles from './Consumerconnectedwebpart.module.scss';
import { IConsumerconnectedwebpartProps } from './IConsumerconnectedwebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList } from 'office-ui-fabric-react';
import { SPOperations } from '../../Services/SPServices';
import { IListItem } from './IListItem';
import { IConsumerconnectedwebpartState } from './IConsumerconnectedwebpartState';
import { IList } from '../../sourceconnectecwebpart/SourceconnectecwebpartWebPart';

export default class Consumerconnectedwebpart extends React.Component<IConsumerconnectedwebpartProps,IConsumerconnectedwebpartState, {}> {
  private _spServices: SPOperations;
  constructor(props:IConsumerconnectedwebpartProps){
    super(props);
    this.state={
      listResults:[],
      ListTitle:""
    };
    this._spServices=new SPOperations(this.props.context);
  }
  public render(): React.ReactElement<IConsumerconnectedwebpartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.consumerconnectedwebpart} ${hasTeamsContext ? styles.teams : ''}`}>
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

  public componentDidUpdate(prevProps: IConsumerconnectedwebpartProps, prevState:IConsumerconnectedwebpartState): void {
    const defaultListTitle: IList ={Title:"CustomList"};
    const dynamicListTitle: IList | string | undefined =this.props.ListTitle.tryGetValue()?this.props.ListTitle.tryGetValue():defaultListTitle.Title;
    if(prevState.ListTitle !==dynamicListTitle.toString() && prevState.ListTitle===this.state.ListTitle){
      this._spServices.getMoreThan5KListItemsRestAPIWithoutWhereClause(this.props.context,dynamicListTitle.toString()).then((results:IListItem[])=>{
        this.setState({listResults:results, ListTitle: dynamicListTitle.toString()});
        console.log("Total results: "+ results.length);
      });
    }
  }

  public componentDidMount(): void {
    const defaultListTitle: IList ={Title:"CustomList"};
    const dynamicListTitle:IList | string | undefined =this.props.ListTitle.tryGetValue()?this.props.ListTitle.tryGetValue():defaultListTitle.Title;
    this._spServices.getMoreThan5KListItemsRestAPIWithoutWhereClause(this.props.context,dynamicListTitle.toString()).then((results:IListItem[])=>{
      this.setState({listResults:results, ListTitle: dynamicListTitle.toString()});
      console.log("Total results: "+ results.length);
    });
  }
}
